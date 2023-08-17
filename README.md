# EntryExitTrack

## 概要

EntryExitTrackは、会社の入退館記録のCSVファイルをExcelに転記するためのバッチ処理プログラムです。

## 使用技術

- 言語: Ruby 3.2.2
- Gem: `rubyXL` (Excelへの書き込み用)

## ファイル構成

- `EntryExitTrack.rb`: バッチ処理のメインプログラムです。CSVファイルの読み込み、データの変換、Excelへの書き込みを行います。

## 使用方法

1. 必要なGemをインストールするために、以下のコマンドを実行してください。

```ruby
gem install rubyXL
```

2. `EntryExitTrack.rb`を実行して、バッチ処理を開始します。

```ruby
ruby EntryExitTrack.rb
```

## 注意事項

- 入退館記録のCSVファイルと転記先のExcelファイルは、`EntryExitTrack.rb`と同じディレクトリに配置してください。
- Excelファイルは、実行後に同じディレクトリに保存されます。
