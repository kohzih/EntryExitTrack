REM ���̃o�b�`�t�@�C���́AEntryExitTrack.rb Ruby�X�N���v�g�����s���A���݂̓��t�Ǝ������t�@�C�����ɒǉ����ꂽ�t�@�C���ɏo�͂����O���܂��B
@echo off
for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set datetime=%%a
set datetime=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%_%datetime:~8,2%%datetime:~10,2%%datetime:~12,2%
ruby EntryExitTrack.rb >> log_%datetime%.txt 2>&1
