require 'csv'
require 'date'
require 'rubyXL'
require 'rubyXL/convenience_methods'

def read_entry_exit_csv
  data = {}
  # "output"で始まるcsvファイルは除いて処理する
  Dir.glob('./*.csv').reject{|f| f =~ /^\.\/output.*\.csv$/}.each do |file|
    CSV.foreach(file, headers: true, skip_lines: /^1/).with_index do |row, idx|
      # next if idx.zero? # 1行目は無視
      next if row['氏名'].nil? || row['氏名'].strip.empty? # "氏名"が空の行は無視

      date = row['日時'].split(' ').first # YYYY/MM/DD の部分だけ取得
      time = row['日時'].split(' ').last  # hh:mm:ss の部分だけ取得

      data[date] ||= {}
      data[date][row['氏名']] ||= { earliest: '23:59:59', latest: '00:00:00' }

      # 機器アドレスの下2桁が"02"なら入館、"03"なら退館と判断
      if row['機器アドレス'][-2..-1] == "02"
        # 一番早い時刻には入館時刻のみを設定
        data[date][row['氏名']][:earliest] = [data[date][row['氏名']][:earliest], time].min
      elsif row['機器アドレス'][-2..-1] == "03"
        # 一番遅い時刻には退館時刻のみを設定
        data[date][row['氏名']][:latest] = [data[date][row['氏名']][:latest], time].max
      end
    end
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

def output_csv(data)
  # 結果の出力(タブ区切りのCSVファイル)
  filename = "output_#{Time.now.strftime('%Y%m%d%H%M%S')}.csv"
  CSV.open(filename, 'wb', col_sep: "\t") do |csv|
    csv << ['名前', '入館日時', '退館日時']
    data.sort.each do |name, dates|
      dates.sort.each do |date, times|
        # 氏名の全角・半角スペースは削除。日付と入館時刻、退館時刻を結合して出力する
        csv << [name.delete("　").delete(" "), date + ' ' + times[:earliest], date + ' ' + times[:latest]]
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
        next if row.cells.nil? || row.cells[0].nil? || row.cells[0].value.nil?
  
        # 日付が一致する行が見つかった場合、F列(インデックス5)に入館時刻、G列(インデックス6)に退館時刻を設定
        cell_date = Date.new(1900, 1, 1) + row.cells[0].value.to_i - 2
        date_value = dates[cell_date.strftime('%Y/%m/%d')]
  
        if date_value
          # 入館時刻を設定
          cell = worksheet.add_cell(row.r - 1, 5)
          # セルの書式を時間形式に設定
          cell.set_number_format('h:mm')
          # セルの値を設定
          cell.change_contents(DateTime.strptime(cell_date.strftime('%Y/%m/%d') + ' ' + date_value[:earliest], "%Y/%m/%d %H:%M:%S"))
  
          # 退館時刻を設定
          cell = worksheet.add_cell(row.r - 1, 6)
          cell.set_number_format('h:mm')
          cell.change_contents(DateTime.strptime(cell_date.strftime('%Y/%m/%d') + ' ' + date_value[:latest], "%Y/%m/%d %H:%M:%S"))

          updated = true
        end
      end
    end
  
    # Excelファイルを上書き保存
    workbook.write(excel_file_path) if updated
  end
end

data = read_entry_exit_csv
output_csv(data)
extry_exit_track(data)
puts '正常終了'
