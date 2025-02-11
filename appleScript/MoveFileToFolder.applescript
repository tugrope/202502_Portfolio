-- �f�t�H���g�t�H���_�̃p�X��ݒ�
set defaultPath to "�����Ƀf�t�H���g�t�@�C���p�X��ݒ�"
set destinationPath to "�����Ƀf�t�H���g�t�H���_�p�X1��ݒ�"
set destinationPath2 to "�����Ƀf�t�H���g�t�H���_�p�X2��ݒ�"

-- �ړ���t�H���_�̑��݊m�F
try
	tell application "Finder"
		if not (exists folder destinationPath as POSIX file) then
			display alert "�t�H���_��������܂���B�T�[�o�[�ڑ����čēx���s���Ă��������B"
			return
		end if
		if not (exists folder destinationPath2 as POSIX file) then
			display alert "�ǉ��ۑ���t�H���_��������܂���B�T�[�o�[�ڑ����čēx���s���Ă��������B"
			return
		end if
	end tell
on error
	display alert "�ۑ���t�H���_��������܂���B�T�[�o�[�ڑ����čēx���s���Ă��������B"
	return
end try

-- �t�H���_�I���_�C�A���O��\��
try
	set selectedFolder to choose folder with prompt "�t�H���_��I�����Ă�������" default location (POSIX file defaultPath)
on error
	display alert "�t�H���_�̑I�����L�����Z������܂����B"
	return
end try

-- �I�����ꂽ�t�H���_���̃t�@�C�����ċA�I�Ɍ���
tell application "Finder"
	try
		set mapFiles to (every file of entire contents of selectedFolder whose name ends with "-�n�}ol.ai")
		set fileCount to count of mapFiles

		-- �������ʂɉ����ăA���[�g��\��
		if fileCount is 0 then
			display alert "�Y���t�@�C����0�Ȃ̂ŉ��������ւ�ŁB"
			return
		else
			set alertResult to display alert ("�Y���t�@�C����" & fileCount & "��") message "�y���Ӂz�����t�@�C���͏㏑�����܂��B�t�H���_�փG�N�X�|�[�g�I" buttons {"�L�����Z��", "OK"} default button "OK" cancel button "�L�����Z��"

			if button returned of alertResult is "OK" then
				-- �t�@�C�����ړ�
				set errorFiles to {}
				repeat with aFile in mapFiles
					try
						move aFile to POSIX file destinationPath with replacing
						move aFile to POSIX file destinationPath2 with replacing
					on error errorMessage
						set end of errorFiles to (name of aFile) & ": " & errorMessage
					end try
				end repeat

				-- �G���[�̗L���ɉ���������
				if (count of errorFiles) > 0 then
					set errorList to join of errorFiles with return
					display alert "�ꕔ�̃t�@�C���̈ړ��Ɏ��s���܂����B" message errorList
				else
					-- �G���[���Ȃ��ꍇ�̂݁A�I���t�H���_���S�~���ֈړ�
					try
						move selectedFolder to trash
						display alert "�G�N�X�|�[�g����"
					on error errorMessage
						display alert "�G�N�X�|�[�g����" message "�������A�I���t�H���_�̃S�~���ւ̈ړ��Ɏ��s���܂����B"
					end try
				end if
			end if
		end if
	on error errorMessage
		display alert "�G���[���������܂����B" message errorMessage
	end try
end tell

-- ���X�g�����p�̃n���h��
on join of theList with delimiter
	set AppleScript's text item delimiters to delimiter
	set theString to theList as string
	set AppleScript's text item delimiters to ""
	return theString
end join
