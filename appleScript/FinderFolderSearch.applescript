 try
	-- フォルダ選択ダイアログを表示
	set targetFolder to choose folder with prompt "検索対象のフォルダーを選択してください："

	-- ファイル名またはフォルダ名を入力するダイアログを表示（縦長に変更）
	set dialogResult to display dialog "検索するフォルダ名を入力してください（改行で区切って複数指定可能）：" default answer "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n" with title "検索キーワード入力" with icon note buttons {"キャンセル", "OK"} default button "OK" cancel button "キャンセル"

	-- キャンセルボタンが押された場合は処理を終了
	if button returned of dialogResult is "キャンセル" then
		return
	end if

	set searchTerms to text returned of dialogResult

	-- 検索開始前にアラート表示
	display dialog "完了ボックス表示まで少し待っててね" buttons {"OK"} default button "OK"

	-- 改行区切りのテキストをリストに変換
	set searchList to paragraphs of searchTerms

	-- Finderのオブジェクト参照を取得
	tell application "Finder"
		-- 指定フォルダ内の第一階層のフォルダを取得
		set folderItems to every folder of targetFolder

		-- ヒットした項目を格納するリスト
		set matchedItems to {}

		-- 各検索語で一致する項目を探す
		repeat with searchTerm in searchList
			repeat with anItem in folderItems
				if name of anItem contains (searchTerm as text) then
					set end of matchedItems to anItem
				end if
			end repeat
		end repeat

		-- マッチした項目を選択
		if (count of matchedItems) > 0 then
			select matchedItems
			display dialog ((count of matchedItems) as text) & " 件のフォルダが選択されました。" buttons {"OK"} default button "OK"
		else
			display dialog "一致するフォルダが見つかりませんでした。" buttons {"OK"} default button "OK"
		end if
	end tell

on error errMsg
	display dialog "エラーが発生しました：" & errMsg buttons {"OK"} default button "OK"
end try
