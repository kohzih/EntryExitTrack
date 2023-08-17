require 'csv'
require 'date'
require 'rubyXL'
require 'rubyXL/convenience_methods'

def entry_exit_track
  data = {}
  # output.csvは除いて処理する
  Dir.glob('./*.csv').reject{|f| f == './output.csv'}.each do |file|
    CSV.foreach(file, headers: true, skip_lines: /^1/).with_index do |row, idx|
      # next if idx.zero? # 1行目は無視
      next if row['氏名'].nil? || row['氏名'].strip.empty? # "氏名"が空の行は無視

      date = row['日時'].split(' ').first # YYYY/MM/DD の部分だけ取得
      time = row['日時'].split(' ').last  # hh:mm:ss の部分だけ取得

      data[date] ||= {}
      data[date][row['氏名']] ||= { earliest: '23:59:59', latest: '00:00:00' }

      # 一番早い時刻と遅い時刻を更新
      data[date][row['氏名']][:earliest] = [data[date][row['氏名']][:earliest], time].min
      data[date][row['氏名']][:latest] = [data[date][row['氏名']][:latest], time].max
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
  CSV.open('output.csv', 'wb', col_sep: "\t") do |csv|
    csv << ['名前', '入館日時', '退館日時']
    data.sort.each do |name, dates|
      dates.sort.each do |date, times|
        # 氏名の全角・半角スペースは削除。日付と入館時刻、退館時刻を結合して出力する
        csv << [name.delete("　").delete(" "), date + ' ' + times[:earliest], date + ' ' + times[:latest]]
      end
    end
  end

  # 結果の出力
  # data.each do |date, names|
  #   puts "日付: #{date}"
  #   names.each do |name, times|
  #     puts "氏名: #{name} -> 最初: #{times[:earliest]}, 最後: #{times[:latest]}"
  #   end
  #   puts '-' * 30
  # end
end

data = entry_exit_track
# output_csv(data)

Dir.glob('./*.xlsx').each do |excel_file_path|

  # Excelファイルを開く
  workbook = RubyXL::Parser.parse(excel_file_path)

  data.sort.each do |name, dates|
    # 対応する名前のシートを探す
    worksheet = workbook[name.delete("　").delete(" ")]
    # シートが存在しない場合は次のループへ
    next unless worksheet

    worksheet.each do |row|
      break if row.nil? || row.r > 40 # とりあえず40行目まで
      next if row.cells.nil? || row.cells[0].nil? || row.cells[0].value.nil?

      cell_date = Date.new(1900, 1, 1) + row.cells[0].value.to_i - 2

      date_value = dates[cell_date.strftime('%Y/%m/%d')]

      if date_value
        # 日付が一致する行が見つかった場合、G列(インデックス6)に退館時刻を設定
        # worksheet.add_cell(row.r + 1, 6, cell_date.strftime('%Y/%m/%d') + ' ' + date_value[:latest])

        # Excelの日付形式での日時の値を計算
        # datetime = DateTime.strptime(cell_date.strftime('%Y/%m/%d') + ' ' + date_value[:latest], "%Y/%m/%d %H:%M:%S")
        # excel_datetime = datetime.ajd - Date.new(1899, 12, 30).ajd

        cell = worksheet.add_cell(row.r - 1, 6)

        # セルの書式を時間形式に設定
        cell.set_number_format('h:mm')

        cell.change_contents(DateTime.strptime(cell_date.strftime('%Y/%m/%d') + ' ' + date_value[:latest], "%Y/%m/%d %H:%M:%S"))
      end
    end
  end

  # Excelファイルを上書き保存
  workbook.write(excel_file_path)
end

# # 実行ファイルのディレクトリを取得
# current_directory = File.dirname(__FILE__)

# # 現在のディレクトリ内のすべてのCSVファイルを取得
# csv_files = Dir.glob(File.join(current_directory, "*.csv"))

# # 各CSVファイルの内容をShift-JISからUTF-8に変換
# csv_files.each do |csv_file|
#   # ファイルをShift-JISとして読み込み
#   content = File.open(csv_file, 'r:Shift_JIS') { |f| f.read }

#   # 文字コードを変換。無効なバイトシーケンスや変換できない文字は「?」に置き換える
#   utf8_content = content.encode('UTF-8', 'Shift_JIS', invalid: :replace, undef: :replace)  

#   # ファイルにUTF-8として上書き保存
#   File.write(csv_file, utf8_content)
# end
