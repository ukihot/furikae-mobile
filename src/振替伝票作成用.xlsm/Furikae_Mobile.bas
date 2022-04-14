Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
    'お約束
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Const HEADER_ROW As String = 8
    
    ' 実行確認
    Dim rc As Long
    rc = MsgBox(Format(DateSerial(Year(Now), Month(Now), 0), "yyyy/mm") & " の集計を開始しますがよろしいですか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        End
    End If

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
        Set clone = ThisWorkbook.Worksheets(department_name).Cells(HEADER_ROW, 3)
        original.Copy clone

        ' 保存せずに閉じる
        Call wb.Close(SaveChanges:=False)


        ' A4セルに部署名を記入
        Range("A4") = department_name
        ' F3セルに実行日を記入
        Range("F2") = Format(Date, "yyyy/mm/dd")
        ' C列にて空白じゃなければ左セルにVLOOKUP数式を挿入
        ce = Cells(Rows.Count, "C").End(xlUp).Row
        Dim i As Integer
        For i = HEADER_ROW To ce
            If Not Cells(i, 3) = "" Then
                Cells(i, 2).Formula = "=VLOOKUP(" & Cells(i, 3).Address & ",PHONE_MST!A:B,2,)"
            End If
        Next
        ' E列にて集計作業
        Dim tam
        tam = Format(DateSerial(Year(Now), Month(Now), 0), "mm")
        Dim tac
        Dim ss As Worksheet
        Set ss = ThisWorkbook.Worksheets("summary")

        Select Case tam
            Case "03"
                tac = 2
            Case "04"
                tac = 3
            Case "05"
                tac = 4
            Case "06"
                tac = 5
            Case "07"
                tac = 6
            Case "08"
                tac = 7
            Case "09"
                tac = 8
            Case "10"
                tac = 9
            Case "11"
                tac = 10
            Case "12"
                tac = 11
            Case "01"
                tac = 12
            Case "02"
                tac = 13
            Case Else
                MsgBox "ちょっと待ちんさい"
                End
        End Select


        de = Cells(Rows.Count, "D").End(xlUp).Row
        Dim j
        Dim zeinuki_total As Long: zeinuki_total = 0
        Dim hikazei_total As Long: hikazei_taial = 0

        For j = HEADER_ROW To de
            If Cells(j, 6) = "対象外" Then
                hikazei_total = hikazei_total + Cells(j, 5)
            ElseIf Not Cells(j, 4) = "小計" Then
                zeinuki_total = zeinuki_total + Cells(j, 5)
            End If
        Next

        ' Summaryに書こう

        Dim p As Integer
        For p = 5 To 18
            If ss.Cells(p, 1) = department_name Then
                ss.Cells(p, tac) = zeinuki_total
                GoTo continue1
            End If
        Next
continue1:
        Dim q As Integer
        For q = 39 To 52
            If ss.Cells(q, 1) = department_name Then
                ss.Cells(q, tac) = hikazei_total
                GoTo continue2
            End If
        Next
continue2:
        ' 合計金額を書く
        Range("F5").Formula = "=summary!B" & p & " + summary!B" & q

        hikazei_total = 0
        zeinuki_total = 0

    Next file

    MsgBox "正常に完了しました．"
    Worksheets(1).Activate


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

