Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
    '����
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Const HEADER_ROW As String = 9

    Dim seikyu_month
    seikyu_month = DateSerial(Year(Now), Month(Now), 0)

    ' ���s�m�F
    Dim rc As Long
    rc = MsgBox(Format(seikyu_month, "yyyy/mm") & " �̏W�v���J�n���܂�����낵���ł����H", vbYesNo + vbQuestion)
    If rc = vbNo Then
        End
    End If

    ' �������ʂŃt�H���_�쐬
    Dim tra_path
    tra_path = ThisWorkbook.path & "\���Ə��ʖ���\" & Format(seikyu_month, "yyyymm")
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If Not objFso.FolderExists(tra_path) Then
        objFso.CreateFolder (tra_path)
    End If

    ' tmp�t�H���_�Ɋi�[����Ă���Excel�̐��������̐��ƈ�v(Loop)
    ' ./fetch_bill/tmp/${������}.xlsx��ǂ�

    Dim tmp_path, fso, file, files
    tmp_path = ThisWorkbook.path & "\fetch_bill\tmp\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(tmp_path).files

    '�t�H���_���̑S�t�@�C���ɂ��ď���
    For Each file In files
        ' �t�@�C�����J���ău�b�N�Ƃ��Ď擾
        Dim wb As Workbook
        Set wb = Workbooks.Open(file)
        Dim ws As Worksheet
        Set ws = wb.Worksheets(1)

        Dim department_name As String: department_name = Left(wb.Name, Len(wb.Name) - 5)

        ' �����V�[�g��template�V�[�g�̃R�s�[�Ƃ��č쐬
        If ExistsSheet(department_name) Then
            ThisWorkbook.Sheets(department_name).Delete
        End If
        ThisWorkbook.Worksheets("template").Copy After:=ThisWorkbook.Worksheets(1)
        ActiveSheet.Name = department_name

        ' �����V�[�g�ɕ���Excel�̓��e��]�L
        ' B��́u���v�v�ȍ~�͕s�v�̂��߁C�u���v�v���L�ڂ��ꂽ�s�������
        ' �߂�ǂ��̂ŃG���[�n���h�����O���Ȃ�
        ' -> 20230711:Softbank�C���{�C�X���x�Ή��ɂ��l���ύX
        ' -> ���v=>���v
        Dim goukei_cell As Range
        Set goukei_cell = ws.Columns("B").Find(What:="���v", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows)

        Dim goukei_row
        goukei_row = goukei_cell.Row

        ' �w�b�_�J������[�d�b�ԍ�, ��������, ������z(�~), �ŋ敪]�̌`�ɂȂ��Ă���̂œ]�L
        ' A2 -> D${goukei_row-1} �͈̔͂�C9�ɃR�s�[
        Dim original As Range
        Dim clone As Range
        Set original = ws.Range(ws.Cells(2, 1), ws.Cells(goukei_row - 1, 4))
        Set clone = ThisWorkbook.Worksheets(department_name).Cells(HEADER_ROW, 3)
        original.Copy clone
        ThisWorkbook.Worksheets(department_name).Range("B8", ThisWorkbook.Worksheets(department_name).Cells(goukei_row + 7, 6)).Borders.LineStyle = xlContinuous

        ' �ۑ������ɕ���
        Call wb.Close(SaveChanges:=False)

        ' A4�Z���ɕ��������L��
        Range("A4") = department_name
        ' F3�Z���Ɏ��s�����L��
        Range("F2") = Format(Date, "yyyy/mm/dd")
        ' C��ɂċ󔒂���Ȃ���΍��Z����VLOOKUP������}��
        ce = Cells(Rows.Count, "C").End(xlUp).Row
        Dim i As Integer
        For i = HEADER_ROW To ce
            If Not Cells(i, 3) = "" Then
                Cells(i, 2).Formula = "=VLOOKUP(" & Cells(i, 3).Address & ",PHONE_MST!A:B,2,)"
            End If
        Next
        ' E��ɂďW�v���
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
                MsgBox "������Ƒ҂��񂳂�"
                End
        End Select

        de = Cells(Rows.Count, "D").End(xlUp).Row
        Dim j
        Dim zeinuki_total As Long: zeinuki_total = 0
        Dim hikazei_total As Long: hikazei_taial = 0

        For j = HEADER_ROW To de
            Dim kingaku As Long: kingaku = Cells(j, 5)
            '�ŋ敪���ΏۊO�̏ꍇ�͔�ېŊz�ɍ��v
            If Cells(j, 6) = "�ΏۊO" Then
                hikazei_total = hikazei_total + kingaku
            '���v�ȊO�͑S�ĐŔ����z�ɏW�v
            ' -> 20230711:Softbank�C���{�C�X���x�Ή��ɂ��l���ύX
            ' -> ���v=>�v
            ElseIf Not Cells(j, 4) = "�v" Then
                If Cells(j, 6) = "�� ��" Then
                    kingaku = kingaku / 1.1
                End If
                zeinuki_total = zeinuki_total + kingaku
            End If
        Next

        ' Summary�ɏ�����

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
        ' ���v���z������
        Range("F5") = ss.Cells(p + 16, tac) + ss.Cells(q, tac)

        hikazei_total = 0
        zeinuki_total = 0

        Call ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=tra_path & "\" & department_name & "-�g�ї�������-" & Format(seikyu_month, "yyyymm") & ".xlsx"
        ActiveWorkbook.Close False

    Next file

    MsgBox "����Ɋ������܂����D"
    Worksheets(1).Activate

End Sub

' Sheets �Ɏw�肵�����O�̃V�[�g�����݂��邩���肷��
Public Function ExistsSheet(ByVal bookName As String)
    Dim ws As Variant
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = LCase(bookName) Then
            ExistsSheet = True ' ���݂���
            Exit Function
        End If
    Next

    ' ���݂��Ȃ�
    ExistsSheet = False
End Function

