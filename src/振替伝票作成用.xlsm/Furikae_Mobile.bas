Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
    'お約束
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Const HEADER_ROW As String = 9

    Dim seikyu_month
    seikyu_month = DateSerial(Year(Now), Month(Now), 0)

    ' 実行確認
    Dim rc As Long
    rc = MsgBox(Format(seikyu_month, "yyyy/mm") & " の集計を開始しますがよろしいですか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        End
    End If

    ' 請求月別でフォルダ作成
    Dim tra_path
    tra_path = ThisWorkbook.path & "\事業所別明細\" & Format(seikyu_month, "yyyymm")
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If Not objFso.FolderExists(tra_path) Then
        objFso.CreateFolder (tra_path)
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
        ' -> 20230711:Softbankインボイス制度対応により様式変更
        ' -> 合計=>小計
        Dim goukei_cell As Range
        Set goukei_cell = ws.Columns("B").Find(What:="小計", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows)

        Dim goukei_row
        goukei_row = goukei_cell.Row

        ' ヘッダカラムが[電話番号, 料金内訳, 内訳金額(円), 税区分]の形になっているので転記
        ' A2 -> D${goukei_row-1} の範囲をC9にコピー
        Dim original As Range
        Dim clone As Range
        Set original = ws.Range(ws.Cells(2, 1), ws.Cells(goukei_row - 1, 4))
        Set clone = ThisWorkbook.Worksheets(department_name).Cells(HEADER_ROW, 3)
        original.Copy clone
        ThisWorkbook.Worksheets(department_name).Range("B8", ThisWorkbook.Worksheets(department_name).Cells(goukei_row + 7, 6)).Borders.LineStyle = xlContinuous

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
        Dim tac
        Dim ss As Worksheet
        Set ss = ThisWorkbook.Worksheets("summary")

        Select Case Format(seikyu_month, "mm")
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
            Dim kingaku As Long: kingaku = Cells(j, 5)
            '税区分が対象外の場合は非課税額に合計
            If Cells(j, 6) = "対象外" Then
                hikazei_total = hikazei_total + kingaku
            '小計以外は全て税抜き額に集計
            ' -> 20230711:Softbankインボイス制度対応により様式変更
            ' -> 小計=>計
            ElseIf Not Cells(j, 4) = "計" Then
                If Cells(j, 6) = "内 税" Then
                    kingaku = kingaku / 1.1
                End If
                zeinuki_total = zeinuki_total + kingaku
            End If
        Next

        ' Summaryに書こう

        Dim p As Integer
        For p = 4 To 16
            If ss.Cells(p, 1) = department_name Then
                ss.Cells(p, tac) = zeinuki_total
                GoTo continue1
            End If
        Next
continue1:
        Dim q As Integer
        For q = 36 To 48
            If ss.Cells(q, 1) = department_name Then
                ss.Cells(q, tac) = hikazei_total
                GoTo continue2
            End If
        Next
continue2:
        ' 合計金額を書く
        Range("F5") = ss.Cells(p + 16, tac) + ss.Cells(q, tac)

        hikazei_total = 0
        zeinuki_total = 0

        Call ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=tra_path & "\" & department_name & "-携帯料金明細-" & Format(seikyu_month, "yyyymm") & ".xlsx"
        ActiveWorkbook.Close False

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

