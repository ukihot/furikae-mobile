Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
    'お約束
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' tmpフォルダに格納されているExcelの数が部署の数と一致(Loop)
    ' ./fetch_bill/tmp/${部署名}.xlsxを読む

    Dim tmp_path, fso, file, files
    tmp_path = ThisWorkbook.path & "\fetch_bill\tmp\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(tmp_path).files

    'フォルダ内の全ファイルについて処理
    For Each file In files
        ' ファイルを開いてブックとして取得
        Dim wb As Workbook
        Set wb = Workbooks.Open(file)
        Dim ws As Worksheet
        Set ws = wb.Worksheets(1)

        Dim department_name As String: department_name = Left(wb.Name, Len(wb.Name) - 5)

        ' 部署シートをtemplateシートのコピーとして作成
        If ExistsSheet(department_name) Then
            ThisWorkbook.Sheets(department_name).Delete
        End If
        ThisWorkbook.Worksheets("template").Copy After:=ThisWorkbook.Worksheets(1)
        ActiveSheet.Name = department_name

        ' 部署シートに部署Excelの内容を転記
        ' B列の「合計」以降は不要のため，「合計」が記載された行数を特定

        ' めんどいのでエラーハンドリングしない
        Dim goukei_cell As Range
        Set goukei_cell = ws.Columns("B").Find(What:="合計", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows)

        Dim goukei_row
        goukei_row = goukei_cell.Row

        ' ヘッダカラムが[電話番号, 料金内訳, 内訳金額(円), 税区分]の形になっているので転記
        ' A2 -> D${goukei_row-1} の範囲をC9にコピー
        Dim original As Range
        Dim clone As Range
        Set original = ws.Range(ws.Cells(2, 1), ws.Cells(goukei_row - 1, 4))
        Set clone = ThisWorkbook.Worksheets(department_name).Range("C8")
        original.Copy clone

        ' 保存せずに閉じる
        Call wb.Close(SaveChanges:=False)


        ' A4セルに部署名を記入
        Range("A4") = department_name
        ' F3セルに実行日を記入
        Range("F2") = Format(Date, "yyyy/mm/dd")
        ' C列にて空白じゃなければ左セルにVLOOKUP数式を挿入

        ' E列にて

    Next file

    MsgBox "正常に完了しました．"


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

