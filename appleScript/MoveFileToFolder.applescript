-- デフォルトフォルダのパスを設定
set defaultPath to "ここにデフォルトファイルパスを設定"
set destinationPath to "ここにデフォルトフォルダパス1を設定"
set destinationPath2 to "ここにデフォルトフォルダパス2を設定"

-- 移動先フォルダの存在確認
try
	tell application "Finder"
		if not (exists folder destinationPath as POSIX file) then
			display alert "フォルダが見つかりません。サーバー接続して再度実行してください。"
			return
		end if
		if not (exists folder destinationPath2 as POSIX file) then
			display alert "追加保存先フォルダが見つかりません。サーバー接続して再度実行してください。"
			return
		end if
	end tell
on error
	display alert "保存先フォルダが見つかりません。サーバー接続して再度実行してください。"
	return
end try

-- フォルダ選択ダイアログを表示
try
	set selectedFolder to choose folder with prompt "フォルダを選択してください" default location (POSIX file defaultPath)
on error
	display alert "フォルダの選択がキャンセルされました。"
	return
end try

-- 選択されたフォルダ内のファイルを再帰的に検索
tell application "Finder"
	try
		set mapFiles to (every file of entire contents of selectedFolder whose name ends with "-地図ol.ai")
		set fileCount to count of mapFiles

		-- 検索結果に応じてアラートを表示
		if fileCount is 0 then
			display alert "該当ファイルは0個なので何もせえへんで。"
			return
		else
			set alertResult to display alert ("該当ファイルは" & fileCount & "個") message "【注意】同名ファイルは上書きします。フォルダへエクスポート！" buttons {"キャンセル", "OK"} default button "OK" cancel button "キャンセル"

			if button returned of alertResult is "OK" then
				-- ファイルを移動
				set errorFiles to {}
				repeat with aFile in mapFiles
					try
						move aFile to POSIX file destinationPath with replacing
						move aFile to POSIX file destinationPath2 with replacing
					on error errorMessage
						set end of errorFiles to (name of aFile) & ": " & errorMessage
					end try
				end repeat

				-- エラーの有無に応じた処理
				if (count of errorFiles) > 0 then
					set errorList to join of errorFiles with return
					display alert "一部のファイルの移動に失敗しました。" message errorList
				else
					-- エラーがない場合のみ、選択フォルダをゴミ箱へ移動
					try
						move selectedFolder to trash
						display alert "エクスポート完了"
					on error errorMessage
						display alert "エクスポート完了" message "ただし、選択フォルダのゴミ箱への移動に失敗しました。"
					end try
				end if
			end if
		end if
	on error errorMessage
		display alert "エラーが発生しました。" message errorMessage
	end try
end tell

-- リスト結合用のハンドラ
on join of theList with delimiter
	set AppleScript's text item delimiters to delimiter
	set theString to theList as string
	set AppleScript's text item delimiters to ""
	return theString
end join
