Sub goaikoSelectFilesAndProcessData()
    Dim script As String
    Dim selectedFiles As String
    Dim fileArray() As String
    Dim i As Integer
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsIntermediate As Worksheet
    Dim wsSpecialCheck As Worksheet
    Dim rowOffset As Integer

    ' ファイルダイアログを開き、複数のファイルのパスを取得するAppleScript
    ' 複数のファイルを選択した場合、各ファイルパスは改行(Chr(10))で区切られて返されます
    script = "set theFiles to choose file with prompt ""インポートするExcelファイルを選択してください"" multiple selections allowed true" & Chr(10) & _
             "set filePaths to {}" & Chr(10) & _
             "repeat with aFile in theFiles" & Chr(10) & _
             "set end of filePaths to POSIX path of aFile & ""@@@@""" & Chr(10) & _
             "end repeat" & Chr(10) & _
             "return filePaths as string"

    On Error Resume Next
    selectedFiles = MacScript(script)
    On Error GoTo 0

    If selectedFiles <> "" Then
        ' ファイルパスを@@@@で分割
        fileArray = Split(selectedFiles, "@@@@")

        ' 選択されたファイルのパスをデバッグウインドウに表示
        Dim filePath As Variant
        For i = LBound(fileArray) To UBound(fileArray) - 1  ' 最後の空要素を除外
            Debug.Print "選択されたファイルパス" & Format(i + 1, "00") & ": " & fileArray(i)
        Next i

        Set wsIntermediate = ThisWorkbook.Sheets("整理中間シート（削除不可）")
        Set wsSpecialCheck = ThisWorkbook.Sheets("特殊確認")

        For i = LBound(fileArray) To UBound(fileArray) - 1  ' 最後の空要素を除外
            ' 各ファイルパスをトリムして開く
            Dim trimmedFilePath As String
            trimmedFilePath = Trim(fileArray(i))

            ' ファイルパスが空でないことを確認
            If trimmedFilePath <> "" Then
                On Error Resume Next
                Set wb = Workbooks.Open(trimmedFilePath)

                If Err.Number <> 0 Then
                    Debug.Print "ファイルを開く際にエラーが発生しました: " & trimmedFilePath
                    Debug.Print "エラー " & Err.Number & ": " & Err.Description
                    Err.Clear
                    GoTo ContinueLoop
                End If
                On Error GoTo 0

                If Not wb Is Nothing Then
                    Set wsInput = wb.Sheets("入力シート")

                    ' データを取得して配置
                    ' 最初の空の行を探す
                    rowOffset = 2
                    Do While Not IsEmpty(wsIntermediate.Cells(rowOffset, "Q").Value)
                        rowOffset = rowOffset + 1
                    Loop

                    ' 整理中間シートへのデータ配置:
                    ' Q列: 指定可能1
                    ' R列: 指定可能2
                    ' T列: 指定可能3
                    ' U列: 指定可能4
                    ' AA列: 指定可能5
                    wsIntermediate.Cells(rowOffset, "Q").Value = wsInput.Range("D7").Value ' 指定可能1
                    wsIntermediate.Cells(rowOffset, "R").Value = wsInput.Range("I7").Value ' 指定可能2
                    wsIntermediate.Cells(rowOffset, "AA").Value = wsInput.Range("D12").Value ' 指定可能5
                    wsIntermediate.Cells(rowOffset, "T").Value = wsInput.Range("D20").Value ' 指定可能3
                    wsIntermediate.Cells(rowOffset, "U").Value = wsInput.Range("D21").Value ' 指定可能4

                    ' 特殊確認シートの最初の空行を探す
                    Dim specialCheckRow As Long
                    specialCheckRow = 2
                    Do While Not IsEmpty(wsSpecialCheck.Cells(specialCheckRow, "A").Value)
                        specialCheckRow = specialCheckRow + 1
                    Loop

                    wsSpecialCheck.Cells(specialCheckRow, "A").Value = wsInput.Range("I7").Value ' 指定可能1
                    wsSpecialCheck.Cells(specialCheckRow, "D").Value = wsInput.Range("D10").Value ' 指定可能2
                    wsSpecialCheck.Cells(specialCheckRow, "C").Value = wsInput.Range("D11").Value ' 指定可能3
                    wsSpecialCheck.Cells(specialCheckRow, "B").Value = wsInput.Range("D13").Value ' 指定可能4

                    wb.Close SaveChanges:=False
                End If
            End If
ContinueLoop:
        Next i

        ' すべての処理が完了した後にアラートを表示
        MsgBox "ご愛顧指示書の読み込み完了", vbInformation
    Else
        Debug.Print "ファイルが選択されなかったか、エラーが発生しました。"
        MsgBox "ファイルが選択されなかったか、エラーが発生しました。", vbExclamation
    End If
End Sub
