# EntryExitTrack

## 概要

EntryExitTrackは、会社の入退館記録のCSVファイルをExcelに転記するためのバッチ処理プログラムです。

## 使用技術

- 言語: Ruby 3.2.2
- Gem: `rubyXL` (Excelへの書き込み用)

## ファイル構成

- `EntryExitTrack.rb`: バッチ処理のメインプログラムです。CSVファイルの読み込み、データの変換、Excelへの書き込みを行います。

## CSVファイルの作成

入退館時間が記録されているCSVファイルは、以下の手作業で作成したCSVファイルを使用します。

1. 「セコム セキュリロックⅡ 履歴収集ソフト」でCSVファイルを出力します。
2. 上記のツールで出力したCSVファイルの1行目は不要なヘッダであるため、削除します。
3. 文字のエンコードをShift-JISからUTF-8に変換して保存します。

## 使用方法

1. 必要なGemをインストールするために、以下のコマンドを実行してください。

```ruby
gem install rubyXL
```

2. `EntryExitTrack.rb`を実行して、バッチ処理を開始します。

```ruby
ruby EntryExitTrack.rb
```

## 出力ファイル

- `output_YYYYMMDDHHMMSS.csv`: 入退館記録のCSVファイルのフォーマットを分かりやすくしたものです。こちらのファイルも同時に出力されます。

## 注意事項

- 入退館記録のCSVファイルと転記先のExcelファイルは、`EntryExitTrack.rb`と同じディレクトリに配置してください。
- Excelファイルは、実行後に同じディレクトリに保存されます。
