Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
'お約束
Application.ScreenUpdating = False

' tmpフォルダに格納されているExcelの数が部署の数と一致(Loop)
' ./fetch_bill/tmp/${部署名}.xlsxを読む

Dim tmp_path, fso, file, files
tmp_path = ThisWorkbook.path & "\fetch_bill\tmp\"
Set fso = CreateObject("Scripting.FileSystemObject")
Set files = fso.GetFolder(tmp_path).files

'フォルダ内の全ファイルについて処理
For Each file In files

    'ファイルを開いてブックとして取得
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    Dim department_name As String: department_name = Left(wb.Name, Len(wb.Name) - 5)

    ' 部署名シートをtemplateシートのコピーとして作成
    If Not ExistsSheet(department_name) Then
        ThisWorkbook.Worksheets("template").Copy After:=ThisWorkbook.Worksheets(1)
        ActiveSheet.Name = department_name
    End If
    '保存せずに閉じる
    Call wb.Close(SaveChanges:=False)

Next file


' ヘッダカラムが[電話番号, 料金内訳, 内訳金額(円), 税区分]の形になっているので転記

' B列の「合計」以降は不要


End Sub


' Sheets に指定した名前のシートが存在するか判定する
Public Function ExistsSheet(ByVal bookName As String)
    Dim ws As Variant
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = LCase(bookName) Then
            ExistsSheet = True ' 存在する
            Exit Function
        End If
    Next

    ' 存在しない
    ExistsSheet = False
End Function

