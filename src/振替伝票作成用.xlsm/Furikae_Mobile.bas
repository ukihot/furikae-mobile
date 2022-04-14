Attribute VB_Name = "Furikae_Mobile"
' @Create 2022/04/14
' @Author Yu Tokunaga
Sub Furikae_Mobile()
'����
Application.ScreenUpdating = False

' tmp�t�H���_�Ɋi�[����Ă���Excel�̐��������̐��ƈ�v(Loop)
' ./fetch_bill/tmp/${������}.xlsx��ǂ�

Dim tmp_path, fso, file, files
tmp_path = ThisWorkbook.path & "\fetch_bill\tmp\"
Set fso = CreateObject("Scripting.FileSystemObject")
Set files = fso.GetFolder(tmp_path).files

'�t�H���_���̑S�t�@�C���ɂ��ď���
For Each file In files

    '�t�@�C�����J���ău�b�N�Ƃ��Ď擾
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    Dim department_name As String: department_name = Left(wb.Name, Len(wb.Name) - 5)

    ' �������V�[�g��template�V�[�g�̃R�s�[�Ƃ��č쐬
    If Not ExistsSheet(department_name) Then
        ThisWorkbook.Worksheets("template").Copy After:=ThisWorkbook.Worksheets(1)
        ActiveSheet.Name = department_name
    End If
    '�ۑ������ɕ���
    Call wb.Close(SaveChanges:=False)

Next file


' �w�b�_�J������[�d�b�ԍ�, ��������, ������z(�~), �ŋ敪]�̌`�ɂȂ��Ă���̂œ]�L

' B��́u���v�v�ȍ~�͕s�v


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

