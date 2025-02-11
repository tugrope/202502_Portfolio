 try
	-- �t�H���_�I���_�C�A���O��\��
	set targetFolder to choose folder with prompt "�����Ώۂ̃t�H���_�[��I�����Ă��������F"

	-- �t�@�C�����܂��̓t�H���_������͂���_�C�A���O��\���i�c���ɕύX�j
	set dialogResult to display dialog "��������t�H���_������͂��Ă��������i���s�ŋ�؂��ĕ����w��\�j�F" default answer "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n" with title "�����L�[���[�h����" with icon note buttons {"�L�����Z��", "OK"} default button "OK" cancel button "�L�����Z��"

	-- �L�����Z���{�^���������ꂽ�ꍇ�͏������I��
	if button returned of dialogResult is "�L�����Z��" then
		return
	end if

	set searchTerms to text returned of dialogResult

	-- �����J�n�O�ɃA���[�g�\��
	display dialog "�����{�b�N�X�\���܂ŏ����҂��ĂĂ�" buttons {"OK"} default button "OK"

	-- ���s��؂�̃e�L�X�g�����X�g�ɕϊ�
	set searchList to paragraphs of searchTerms

	-- Finder�̃I�u�W�F�N�g�Q�Ƃ��擾
	tell application "Finder"
		-- �w��t�H���_���̑��K�w�̃t�H���_���擾
		set folderItems to every folder of targetFolder

		-- �q�b�g�������ڂ��i�[���郊�X�g
		set matchedItems to {}

		-- �e������ň�v���鍀�ڂ�T��
		repeat with searchTerm in searchList
			repeat with anItem in folderItems
				if name of anItem contains (searchTerm as text) then
					set end of matchedItems to anItem
				end if
			end repeat
		end repeat

		-- �}�b�`�������ڂ�I��
		if (count of matchedItems) > 0 then
			select matchedItems
			display dialog ((count of matchedItems) as text) & " ���̃t�H���_���I������܂����B" buttons {"OK"} default button "OK"
		else
			display dialog "��v����t�H���_��������܂���ł����B" buttons {"OK"} default button "OK"
		end if
	end tell

on error errMsg
	display dialog "�G���[���������܂����F" & errMsg buttons {"OK"} default button "OK"
end try
