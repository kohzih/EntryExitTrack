REM このバッチファイルは、EntryExitTrack.rb Rubyスクリプトを実行し、現在の日付と時刻がファイル名に追加されたファイルに出力をログします。
@echo off
for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set datetime=%%a
set datetime=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%%datetime:~8,2%%datetime:~10,2%%datetime:~12,2%
ruby EntryExitTrack.rb >> log_%datetime%.txt 2>&1
