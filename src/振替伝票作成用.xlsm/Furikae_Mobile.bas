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

    'ブックに対する処理

    '保存せずに閉じる
    Call wb.Close(SaveChanges:=False)

Next file
' 部署名シートをtemplateシートのコピーとして作成





' ヘッダカラムが[電話番号, 料金内訳, 内訳金額(円), 税区分]の形になっているので転記

' B列の「合計」以降は不要


End Sub
