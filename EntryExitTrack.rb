require 'csv'
require 'date'
require 'rubyXL'
require 'rubyXL/convenience_methods'

DATE_CELL_INDEX = 0
ENTRY_TIME_CELL_INDEX = 5
EXIT_TIME_CELL_INDEX = 6
# 機器アドレスの下2桁が"02"なら入館、"03"なら退館
ENTRY_CODE = '02'
EXIT_CODE = '03'

def read_entry_exit_csv
  data = {}

  # CSVファイルを、Shift-JISからUTF-8に変換し、結果をファイル名の先頭に"input_"を付加したCSVファイルに保存する
  Dir.glob('./*.csv').each do |file|
    # ファイル名の先頭が"input_"または"summary_"で始まるファイルは処理対象外とする
    file_name = File.basename(file)  # ファイルの名前だけを取得
    next if file_name.start_with?('input_', 'summary_')

    convert_and_trim_file(file)
  end

  # 上記処理で作成した、"input_"で始まるcsvファイルのみ処理対象とする
  Dir.glob('./input_*.csv').each do |file|
    CSV.foreach(file, headers: true, skip_lines: /^1/).with_index do |row, idx|
      # 1行目または、"氏名"が空の行は無視する
      next if row['氏名'].nil? || row['氏名'].strip.empty?

      date = row['日時'].split(' ').first # YYYY/MM/DD の部分だけ取得
      time = row['日時'].split(' ').last  # hh:mm:ss の部分だけ取得

      data[date] ||= {}
      data[date][row['氏名']] ||= { earliest: '__:__:__', latest: '__:__:__' }

      # 機器アドレスの下2桁で入館、退館を判断
      if row['機器アドレス'][-2..-1] == ENTRY_CODE
        # 一番早い時刻には入館時刻のみを設定
        earliest_time = data[date][row['氏名']][:earliest]
        data[date][row['氏名']][:earliest] = earliest_time == '__:__:__' ? time : [earliest_time, time].min
      elsif row['機器アドレス'][-2..-1] == EXIT_CODE
        # 一番遅い時刻には退館時刻のみを設定
        latest_time = data[date][row['氏名']][:latest]
        data[date][row['氏名']][:latest] = latest_time == '__:__:__' ? time : [latest_time, time].max
      end
    end
    # 読み込んだCSVファイルを削除する
    File.delete(file)
  end

  # 日付単位でキャッシュされているデータを氏名単位でキャッシュする
  new_data = {}
  data.each do |date, names|
    names.each do |name, times|
      new_data[name] ||= {}
      new_data[name][date] = times
    end
  end

  data = new_data

  return data
end

def convert_and_trim_file(input_file_path)
  # ファイル名の先頭に"input_"を付加する
  output_file_path = input_file_path.sub(/(\w+)\.csv$/, 'input_\1.csv')

  # ファイルをShift-JISエンコーディングで読み込む
  lines = File.readlines(input_file_path, encoding: 'Shift_JIS')
  lines.shift # 先頭の1行を削除する

  # Shift_JISからUTF-8に直接変換できない文字を個別に置換する
  lines.map! do |line|
    line.encode("UTF-8", fallback: {
      "\xFBM".force_encoding('Shift_JIS') => "濵",
      "\xFA\xBA".force_encoding('Shift_JIS') => "德"
    })
  end

  # ファイルをUTF-8エンコーディングで保存し直す
  File.open(output_file_path, 'w', encoding: 'UTF-8') do |f|
    f.puts(lines)
  end
end

def output_summary_text(data)
  # 要約したファイルの出力(タブ区切りのCSVファイル)
  filename = "summary_#{Time.now.strftime('%Y%m%d_%H%M%S')}.txt"
  CSV.open(filename, 'wb', col_sep: "\t") do |csv|
    csv << ['氏名'.ljust(7, '　'), '入館日時', '退館日時']
    data.sort.each do |name, dates|
      dates.sort.each do |date, times|
        # 氏名の全角・半角スペースは削除。日付と入館時刻、退館時刻を結合して出力する
        csv << [name.delete("　").delete(" ").ljust(7, '　'), date + ' ' + times[:earliest], date + ' ' + times[:latest]]
      end
    end
  end
end

def extry_exit_track(data)
  Dir.glob('./*.xlsx').each do |excel_file_path|
    # Excelファイルを開く
    workbook = RubyXL::Parser.parse(excel_file_path)
    updated = false

    data.sort.each do |name, dates|
      # 対応する名前のシートを探す
      worksheet = workbook[name.delete("　").delete(" ")]
      # シートが存在しない場合は次のループへ
      next unless worksheet

      worksheet.each do |row|
        break if row.nil? || row.r > 40 # とりあえず40行目まで
        next if row.cells.nil? || row.cells[DATE_CELL_INDEX].nil? || row.cells[DATE_CELL_INDEX].value.nil?

        # 日付が一致する行が見つかった場合、F列(インデックス5)に入館時刻、G列(インデックス6)に退館時刻を設定
        cell_date = Date.new(1900, 1, 1) + row.cells[DATE_CELL_INDEX].value.to_i - 2
        date_value = dates[cell_date.strftime('%Y/%m/%d')]

        next if date_value.nil? # 日付が一致するデータがない場合は次のループへ

        # 入館時刻を設定
        if date_value[:earliest] != '__:__:__'
          set_time_in_cell(worksheet, row, ENTRY_TIME_CELL_INDEX, date_value[:earliest], cell_date)
          updated = true
        end
        # 退館時刻を設定
        if date_value[:latest] != '__:__:__'
          set_time_in_cell(worksheet, row, EXIT_TIME_CELL_INDEX, date_value[:latest], cell_date)
          updated = true
        end
      end
    end

    # Excelファイルを上書き保存
    workbook.write(excel_file_path) if updated
  end
end

def set_time_in_cell(worksheet, row, cell_index, date_value, cell_date)
  cell = worksheet.add_cell(row.r - 1, cell_index)
  cell.set_number_format('h:mm')
  cell.change_contents(DateTime.strptime(cell_date.strftime('%Y/%m/%d') + ' ' + date_value, "%Y/%m/%d %H:%M:%S"))
end

data = read_entry_exit_csv
output_summary_text(data)
extry_exit_track(data)
puts '正常終了'
