Option Explicit

' �萔��`

' ���o�����֘A�萔
Private Const ���o���J�n�s As Long = 51 ' ���o���f�[�^�J�n�s
Private Const ���o������ As Long = 2     ' B��F���o����
Private Const �E�v�� As Long = 3         ' C��F�E�v
Private Const ���o�����z�� As Long = 4   ' D��F���o�����z
Private Const �c���� As Long = 5         ' E��F�c��
Private Const ���ԍό����� As Long = 6 ' F��F���ؒ��̖��ԍό������v

' �������֘A�萔
Private Const �������� As Long = 3  ' C��
Private Const �������s As Long = 30  ' 30�s��

' �ؓ������֘A�萔
Private Const �ؓ������� As Long = 2  ' B��F�ؓ�����
Private Const �ؓ������J�n���� As Long = 3  ' C��F�J�n��
Private Const �ؓ������J�n�s As Long = 34  ' 24�s��

' �x�����Q�������֘A�萔
Private Const �x�����Q�������� As Long = 2  ' B��F�x�����Q������
Private Const �x�����Q�������J�n���� As Long = 3  ' C��F�J�n��
Private Const �x�����Q�������J�n�s As Long = 15  ' 15�s��

' �v�Z���쐬�p�X�֘A�萔
Private Const �v�Z���쐬�p�X�� As Long = 3  ' C��F�v�Z���쐬�p�X
Private Const �v�Z���쐬�p�X�s As Long = 7  ' 7�s��

' �ڋq�ԍ��֘A�萔
Private Const �ڋq�ԍ��� As Long = 3  ' C��F�ڋq�ԍ�
Private Const �ڋq�ԍ��s As Long = 6  ' 6�s��

' �葱���R�֘A�萔
Private Const �葱���R�� As Long = 3  ' C��F�葱���R
Private Const �葱���R�s As Long = 10  ' 10�s��

' �葱�J�n���֘A�萔
Private Const �葱�J�n���� As Long = 3  ' C��F�葱�J�n��
Private Const �葱�J�n���s As Long = 11  ' 11�s��

' ���[�������X�e�[�^�X�֘A�萔
Private Const ���[�������X�e�[�^�X�� As Long = 3  ' C��F���[�������X�e�[�^�X
Private Const ���[�������X�e�[�^�X�s As Long = 27  ' 27�s��

' �������R�֘A�萔
Private Const �������R�� As Long = 5  ' E��F�������R
Private Const �������R�s As Long = 30  ' 30�s��

' �E�v������֘A�萔
Private Const �ؓ��E�v�ؓ������� As String = "�ؓ�"     ' �ؓ��������E�v������
Private Const �ؓ��E�v�؊������� As String = "�؊�"     ' �؊��������E�v������
Private Const �ԍϓE�v�ԍϕ������� As String = "�ԍϕ�"   ' �ԍς������E�v������

' ���[�������X�e�[�^�X�֘A�萔
Private Const �����X�e�[�^�X������ As String = "����"  ' �����������X�e�[�^�X������
Private Const �����؂ꗝ�R������ As String = "�����؂�"  ' �����؂���������R������
Private Const ����X�e�[�^�X������ As String = "����"  ' ����������X�e�[�^�X������
Private Const ���ԍσC�x���g������ As String = "���ԍ�"  ' ���ԍς������C�x���g������
Private Const �������X�e�[�^�X������ As String = "�����i���j"  ' �����i���j�������X�e�[�^�X������
Private Const ���؃C�x���g������ As String = "����"  ' ���؂������C�x���g������
Private Const �����C�x���g������ As String = "����"  ' �����������C�x���g������
Private Const �j�Y�C�x���g������ As String = "�j�Y"  ' �j�Y�C�x���g������
Private Const �����E�v������ As String = "����"  ' �����������E�v������
Private Const �x�����Q���E�v������ As String = "�x�����Q��"  ' �x�����Q���������E�v������

' ���t�֘A�萔
Private Const ���t�����l As Date = #1/1/1900#  ' �󔒓��t�̏����l

' ���[�N�V�[�g���֘A�萔
Private Const �c�[���V�[�g�� As String = "�c�[��"  ' �c�[���V�[�g�̖��O
Private Const �e���v���[�g�V�[�g�� As String = "�e���v���[�g_EXCEL"  ' �e���v���[�g�V�[�g�̖��O

' �o�͊֘A�萔
Private Const �o�͊J�n�s�I�t�Z�b�g As Long = 8  ' A9�Z������\��t���邽�߂̃I�t�Z�b�g
Private Const �o�͌ڋq�ԍ��s As Long = 4  ' B4�s�F�ڋq�ԍ�
Private Const �o�͌ڋq�ԍ��� As Long = 2  ' B��F�ڋq�ԍ�
Private Const �o�͎葱�J�n���s As Long = 2  ' J2�s�F�葱�J�n��
Private Const �o�͎葱�J�n���� As Long = 10  ' J��F�葱�J�n��
Private Const �o�͊������s As Long = 3  ' J3�s�F������
Private Const �o�͊������� As Long = 10  ' J��F������
Private Const �o�͊������R�s As Long = 3  ' K3�s�F�������R
Private Const �o�͊������R�� As Long = 11  ' K��F�������R

' �o�͍��ڗ�萔�iA��`S��j
Private Const �o��_�ʔԗ� As Long = 1          ' A��F�ʔ�
Private Const �o��_�X�e�[�^�X�� As Long = 2      ' B��F�X�e�[�^�X
Private Const �o��_�C�x���g�� As Long = 3        ' C��F�C�x���g
Private Const �o��_���ԍό��� As Long = 4      ' D��F���ԍό�
Private Const �o��_�Ώی����� As Long = 5        ' E��F�Ώی���
Private Const �o��_�v�Z���ԊJ�n���� As Long = 6  ' F��F�v�Z���ԊJ�n��
Private Const �o��_��؂�� As Long = 7          ' G��F"�`"
Private Const �o��_�v�Z���ԏI������ As Long = 8  ' H��F�v�Z���ԏI����
Private Const �o��_�v�Z������ As Long = 9        ' I��F�v�Z����
Private Const �o��_������ As Long = 10           ' J��F����
Private Const �o��_�ϐ��� As Long = 11           ' K��F�ϐ�
Private Const �o��_�������z�� As Long = 12       ' L��F�������z
Private Const �o��_�x�����Q���� As Long = 13     ' M��F�x�����Q��
Private Const �o��_�ؓ����� As Long = 14         ' N��F�ؓ���
Private Const �o��_�ؓ��z�� As Long = 15         ' O��F�ؓ��z
Private Const �o��_�ԍϓ��� As Long = 16         ' P��F�ԍϓ�
Private Const �o��_����_�ԍϊz�� As Long = 17    ' Q��F����_�ԍϊz
Private Const �o��_����_�ԍϊz�� As Long = 18    ' R��F����_�ԍϊz
Private Const �o��_�x����_�ԍϊz�� As Long = 19  ' S��F�x����_�ԍϊz

' �ԍϗ\����̒萔
Const �ԍϗ\��J�n�s As Long = 40  ' 40�s��
Const �ԍϗ\����� As Long = 3  ' C��
Const �ԍό����� As Long = 4    ' D��

' �ԍϗ������̒萔
Const �ԍϗ����J�n�s As Long = 75   ' 75�s��
Const �ԍϗ�����t�� As Long = 2    ' B��F���t
Const �ԍϗ���E�v�� As Long = 3    ' C��F�E�v
Const �ԍϗ����o�����z�� As Long = 4 ' D��F�o�����z

' �폜�Ō�s�ڂ̒萔
Const �폜�Ō�s�� As Long = 69

' �f�[�^�\��t���J�n�s�̒萔
Const �f�[�^�\��t���J�n�s As Long = �o�͊J�n�s�I�t�Z�b�g + 1

' �������R�擾�֐�
' E��25�s�ڂ���������R���擾���A���[�������X�e�[�^�X�ɉ����ď����𕪊�
Public Function �������R�擾(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    Dim ���[�������X�e�[�^�X As String
    
    ' E��25�s�ڂ̒l���擾
    cellValue = targetSheet.Cells(�������R�s, �������R��).Value
    
    ' ���[�������X�e�[�^�X���擾
    ���[�������X�e�[�^�X = ���[�������X�e�[�^�X�擾(targetSheet)
    
    ' �󔒃`�F�b�N
    If cellValue = "" Or IsEmpty(cellValue) Then
        ' ���[�������X�e�[�^�X���u�����v�̏ꍇ�͓��͕K�{
        If ���[�������X�e�[�^�X = �����X�e�[�^�X������ Then
            Err.Raise 13, "�������R�擾", "�������R���ݒ肳��Ă��܂���BE25�Z���Ɋ������R����͂��Ă��������B"
        Else
            ' ����ȊO�̏ꍇ�́u�����؂�v��ݒ�
            �������R�擾 = �����؂ꗝ�R������
            Exit Function
        End If
    End If
    
    ' ������Ƃ��ĕϊ�
    �������R�擾 = CStr(cellValue)
End Function

' ���[�������X�e�[�^�X�擾�֐�
' C��22�s�ڂ��烍�[�������X�e�[�^�X���擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
Public Function ���[�������X�e�[�^�X�擾(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C��22�s�ڂ̒l���擾
    cellValue = targetSheet.Cells(���[�������X�e�[�^�X�s, ���[�������X�e�[�^�X��).Value
    
    ' �󔒃`�F�b�N
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "���[�������X�e�[�^�X�擾", "���[�������X�e�[�^�X���ݒ肳��Ă��܂���BC22�Z���Ƀ��[�������X�e�[�^�X����͂��Ă��������B"
    End If
    
    ' ������Ƃ��ĕϊ�
    ���[�������X�e�[�^�X�擾 = CStr(cellValue)
End Function

' �葱�J�n���擾�֐�
' C��11�s�ڂ���葱�J�n�����擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
Public Function �葱�J�n���擾(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' C��11�s�ڂ̒l���擾
    cellValue = targetSheet.Cells(�葱�J�n���s, �葱�J�n����).Value
    
    ' �󔒃`�F�b�N
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "�葱�J�n���擾", "�葱�J�n�����ݒ肳��Ă��܂���BC11�Z���Ɏ葱�J�n������͂��Ă��������B"
    End If
    
    ' ���t�^�`�F�b�N
    If Not IsDate(cellValue) Then
        Err.Raise 13, "�葱�J�n���擾", "�葱�J�n�������t�ł͂���܂���BC11�Z���ɐ��������t����͂��Ă��������B"
    End If
    
    ' ���t�^�Ƃ��ĕϊ�
    �葱�J�n���擾 = CDate(cellValue)
End Function

' �葱���R�擾�֐�
' C��10�s�ڂ���葱���R���擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
Public Function �葱���R�擾(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C��10�s�ڂ̒l���擾
    cellValue = targetSheet.Cells(�葱���R�s, �葱���R��).Value
    
    ' �󔒃`�F�b�N
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "�葱���R�擾", "�葱���R���ݒ肳��Ă��܂���BC10�Z���Ɏ葱���R����͂��Ă��������B"
    End If
    
    ' ������Ƃ��ĕϊ�
    �葱���R�擾 = CStr(cellValue)
End Function

' �ڋq�ԍ��擾�֐�
' C��6�s�ڂ���ڋq�ԍ����擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
Public Function �ڋq�ԍ��擾(targetSheet As Worksheet) As String
    Dim cellValue As Variant
    
    ' C��6�s�ڂ̒l���擾
    cellValue = targetSheet.Cells(�ڋq�ԍ��s, �ڋq�ԍ���).Value
    
    ' �󔒃`�F�b�N
    If cellValue = "" Or IsEmpty(cellValue) Then
        Err.Raise 13, "�ڋq�ԍ��擾", "�ڋq�ԍ����ݒ肳��Ă��܂���BC6�Z���Ɍڋq�ԍ�����͂��Ă��������B"
    End If
    
    ' ������Ƃ��ĕϊ�
    �ڋq�ԍ��擾 = CStr(cellValue)
End Function

' �v�Z���̍쐬�p�X�擾�֐�
' C��7�s�ڂ���v�Z���̍쐬�p�X���擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
' �p�X�����݂���t�H���_�łȂ��ꍇ���G���[�𔭐�������
Public Function �v�Z���̍쐬�p�X�擾(targetSheet As Worksheet) As String
    Dim pathValue As Variant
    Dim pathString As String
    
    ' C��7�s�ڂ̒l���擾
    pathValue = targetSheet.Cells(�v�Z���쐬�p�X�s, �v�Z���쐬�p�X��).Value
    
    ' �󔒃`�F�b�N
    If pathValue = "" Or IsEmpty(pathValue) Then
        Err.Raise 13, "�v�Z���̍쐬�p�X�擾", "C��7�s�ځi�v�Z���̍쐬�p�X�j���󔒂ł��B"
    End If
    
    ' ������Ƃ��ĕϊ�
    pathString = CStr(pathValue)
    
    ' �p�X�̑Ó����`�F�b�N�i�t�H���_�����݂��邩�`�F�b�N�j
    If Dir(pathString, vbDirectory) = "" Then
        Err.Raise 76, "�v�Z���̍쐬�p�X�擾", "�w�肳�ꂽ�p�X '" & pathString & "' �̓t�H���_�ł͂���܂���B"
    End If
    
    ' ������Ƃ��ĕԂ�
    �v�Z���̍쐬�p�X�擾 = pathString
End Function

' �t�@�C���o�͊֐�
' �v�Z���̍쐬�p�X�擾�Ŏ擾�����t�H���_�ɗ����v�Z���t�@�C�����쐬���A���핪�o�̓f�[�^��\��t����
Public Sub �t�@�C���o��(targetSheet As Worksheet, templateSheet As Worksheet)
    Dim �o�̓t�H���_�p�X As String
    Dim �o�̓f�[�^ As Variant
    Dim �V�������[�N�u�b�N As Workbook
    Dim �V�������[�N�V�[�g As Worksheet
    Dim �t�@�C���� As String
    Dim ���S�t�@�C���p�X As String
    Dim ���ݓ��� As Date
    Dim �N���������� As String
    Dim �����b������ As String
    
    On Error GoTo ErrorHandler
    
    ' 1. �v�Z���̍쐬�p�X�擾
    �o�̓t�H���_�p�X = �v�Z���̍쐬�p�X�擾(targetSheet)
    
    ' 2. �o�̓f�[�^�쐬
    �o�̓f�[�^ = �o�̓f�[�^�쐬(targetSheet)
    
    ' 3. ���ݓ������擾���ăt�@�C�������쐬
    ���ݓ��� = Now
    �N���������� = Format(���ݓ���, "yyyymmdd")
    �����b������ = Format(���ݓ���, "hhmmss")
    Dim �ڋq�ԍ� As String
    �ڋq�ԍ� = �ڋq�ԍ��擾(targetSheet)
    �t�@�C���� = "�����v�Z��" & �ڋq�ԍ� & ".xlsx"
    
    ' 4. ���S�t�@�C���p�X���쐬
    If Right(�o�̓t�H���_�p�X, 1) <> "\" Then
        ���S�t�@�C���p�X = �o�̓t�H���_�p�X & "\" & �t�@�C����
    Else
        ���S�t�@�C���p�X = �o�̓t�H���_�p�X & �t�@�C����
    End If
    
    ' 5. �V�������[�N�u�b�N���쐬
    Set �V�������[�N�u�b�N = Workbooks.Add
    
    ' 6. �e���v���[�g�V�[�g��V�������[�N�u�b�N�ɃR�s�[
    ' ��\���V�[�g�̏ꍇ�A�ꎞ�I�ɕ\�����Ă���R�s�[
    Dim ���̕\����� As XlSheetVisibility
    ���̕\����� = templateSheet.Visible
    If templateSheet.Visible <> xlSheetVisible Then
        templateSheet.Visible = xlSheetVisible
    End If
    
    templateSheet.Copy Before:=�V�������[�N�u�b�N.Worksheets(1)
    
    ' ���̕\����Ԃɖ߂�
    templateSheet.Visible = ���̕\�����
    Set �V�������[�N�V�[�g = �V�������[�N�u�b�N.Worksheets(1)
    �V�������[�N�V�[�g.Name = "�����v�Z��"
    
    ' ����Sheet1���폜
    Application.DisplayAlerts = False
    �V�������[�N�u�b�N.Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    ' 7. �f�[�^��A9�Z������\��t��
    If IsArray(�o�̓f�[�^) Then
        Dim �s�� As Long
        Dim �� As Long
        �s�� = UBound(�o�̓f�[�^, 1)
        �� = UBound(�o�̓f�[�^, 2)
        
        ' ��ʍX�V���~���ăp�t�H�[�}���X������
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        ' �s�v�ȍs���폜�i�R�s�[�����Ō�̍s����폜�Ō�s�ڂ܂ō폜�j
        Dim �Ō�̍s As Long
        �Ō�̍s = �f�[�^�\��t���J�n�s + �s�� - 2  ' �f�[�^�\��t���J�n�s����\��t�����Ō�̍s(�u�v�v�s����)
        Dim �폜�J�n�s As Long
        �폜�J�n�s = �Ō�̍s + 1
        
        ' �폜�͈͂����݂���ꍇ�̂ݍ폜
        If �폜�J�n�s <= �폜�Ō�s�� Then
            �V�������[�N�V�[�g.Rows(�폜�J�n�s & ":" & �폜�Ō�s��).Delete
        End If

        ' �������������l�����ăZ���͈͂��w�肵�ē\��t��
        Dim �\��t���͈� As Range
        Set �\��t���͈� = �V�������[�N�V�[�g.Range("A" & �f�[�^�\��t���J�n�s).Resize(�s��, ��)
        
        ' �f�[�^��\��t��
        �\��t���͈�.Value = �o�̓f�[�^
        
        ' ��ʍX�V���ĊJ
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
    
    ' 7.5. �ڋq�ԍ��ݒ�
    �V�������[�N�V�[�g.Cells(�o�͌ڋq�ԍ��s, �o�͌ڋq�ԍ���).Value = "�ڋq�ԍ�" & �ڋq�ԍ�
    
    ' 7.6. �葱�J�n���ݒ�
    Dim �葱�J�n�� As Date
    �葱�J�n�� = �葱�J�n���擾(targetSheet)
    �V�������[�N�V�[�g.Cells(�o�͎葱�J�n���s, �o�͎葱�J�n����).Value = �葱�J�n��
    
    ' 7.7. �������ݒ�
    Dim ������ As Date
    ������ = �������擾(targetSheet)
    �V�������[�N�V�[�g.Cells(�o�͊������s, �o�͊�������).Value = ������
    
    ' 7.8. �������R�ݒ�
    Dim �������R As String
    �������R = �������R�擾(targetSheet)
    �V�������[�N�V�[�g.Cells(�o�͊������R�s, �o�͊������R��).Value = �������R
    
    ' 8. �t�@�C���ۑ��i�����t�@�C��������ꍇ�͘A�ԕt���ŕۑ��j
    Dim �ۑ��t�@�C���p�X As String
    Dim �J�E���^ As Long
    �ۑ��t�@�C���p�X = ���S�t�@�C���p�X
    �J�E���^ = 1
    
    ' �����t�@�C��������ꍇ�͘A�Ԃ�t���ĐV�����t�@�C�������쐬
    Do While Dir(�ۑ��t�@�C���p�X) <> ""
        Dim �t�@�C�������� As String
        Dim �g���q���� As String
        Dim �t�H���_�p�X���� As String
        
        ' �t�@�C���p�X�𕪉�
        �t�H���_�p�X���� = Left(���S�t�@�C���p�X, InStrRev(���S�t�@�C���p�X, "\"))
        �t�@�C�������� = Mid(���S�t�@�C���p�X, InStrRev(���S�t�@�C���p�X, "\") + 1)
        �g���q���� = Right(�t�@�C��������, 5) ' ".xlsx"
        �t�@�C�������� = Left(�t�@�C��������, Len(�t�@�C��������) - 5)
        
        ' �A�ԕt���t�@�C�������쐬
        �ۑ��t�@�C���p�X = �t�H���_�p�X���� & �t�@�C�������� & "(" & �J�E���^ & ")" & �g���q����
        �J�E���^ = �J�E���^ + 1
    Loop
    
    �V�������[�N�u�b�N.SaveAs Filename:=�ۑ��t�@�C���p�X, FileFormat:=xlOpenXMLWorkbook
    
    ' 9. ���[�N�u�b�N�����
    �V�������[�N�u�b�N.Close SaveChanges:=False
    
    ' 10. �������b�Z�[�W
    MsgBox "�����v�Z���t�@�C���̏o�͂��������܂����B" & vbCrLf & "�ۑ���: " & �ۑ��t�@�C���p�X, vbInformation, "�t�@�C���o�͊���"
    
    Exit Sub
    
ErrorHandler:
    ' �G���[�����������ꍇ�̓��[�N�u�b�N�����
    If Not �V�������[�N�u�b�N Is Nothing Then
        �V�������[�N�u�b�N.Close SaveChanges:=False
    End If
    
    ' �G���[���b�Z�[�W��\��
    MsgBox "�t�@�C���o�͒��ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
    Err.Raise Err.Number, "�t�@�C���o��", Err.Description
End Sub

' �v�Z���쐬���C������
' �c�[���V�[�g��ΏۂƂ��ăt�@�C���o�͂����s����
Public Sub �v�Z���쐬()
    Dim �c�[���V�[�g As Worksheet
    Dim �e���v���[�g�V�[�g As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' �c�[���V�[�g���擾
    Set �c�[���V�[�g = ThisWorkbook.Worksheets(�c�[���V�[�g��)
    
    ' �e���v���[�g�V�[�g���擾
    Set �e���v���[�g�V�[�g = ThisWorkbook.Worksheets(�e���v���[�g�V�[�g��)
    
    ' �t�@�C���o�͂����s
    Call �t�@�C���o��(�c�[���V�[�g, �e���v���[�g�V�[�g)
    
    Exit Sub
    
ErrorHandler:
    'MsgBox "�v�Z���쐬���ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' �ԍϗ\����擾�֐�
' 35�s�ڂ���J�n���AC�񂪋󔒂ɂȂ�܂ŕԍϗ\����A�ԍό����A�ԍό����݌v���擾
Public Function �ԍϗ\����擾(targetSheet As Worksheet) As Variant
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim �ԍό����݌v As Double
    
    ' 35�s�ڂ�C�񂪋󔒂̏ꍇ��1�s�Ƃ��ď���
    currentRow = �ԍϗ\��J�n�s + 1
    rowCount = 1
    
    ' �f�[�^�s�����J�E���g�iC�񂪋󔒂ɂȂ�܂Łj
    Do While targetSheet.Cells(currentRow, �ԍϗ\�����).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
        
    ' �z����������i�s�� x 3��F�ԍϗ\����A�ԍό����A�ԍό����݌v�j
    ReDim dataArray(1 To rowCount, 1 To 3)
    
    ' �f�[�^���擾���ăo���f�[�V����
    �ԍό����݌v = 0
    currentRow = �ԍϗ\��J�n�s
    
    For i = 1 To rowCount
        ' C��F�ԍϗ\���
        Dim dateValue As Variant
        dateValue = targetSheet.Cells(currentRow, �ԍϗ\�����).Value
        
        If dateValue = "" Or IsEmpty(dateValue) Then
            ' �󔒂̏ꍇ�͓��t�̏����l��ݒ�
            dataArray(i, 1) = ���t�����l
        Else
            ' ���t�^�`�F�b�N
            If Not IsDate(dateValue) Then
                Err.Raise 13, "�ԍϗ\����擾", currentRow & "�s�ڂ�C��i�ԍϗ\����j�����t�ł͂���܂���B"
            End If
            dataArray(i, 1) = CDate(dateValue)
        End If
        
        ' D��F�ԍό���
        Dim principalValue As Variant
        principalValue = targetSheet.Cells(currentRow, �ԍό�����).Value
        
        If principalValue = "" Or IsEmpty(principalValue) Then
            ' �󔒂̏ꍇ��0��ݒ�
            dataArray(i, 2) = 0
        Else
            ' ���l�^�`�F�b�N
            If Not IsNumeric(principalValue) Then
                Err.Raise 13, "�ԍϗ\����擾", currentRow & "�s�ڂ�D��i�ԍό����j�����l�ł͂���܂���B"
            End If
            dataArray(i, 2) = CDbl(principalValue)
        End If
        
        ' �ԍό����݌v���v�Z
        �ԍό����݌v = �ԍό����݌v + dataArray(i, 2)
        dataArray(i, 3) = �ԍό����݌v
        
        currentRow = currentRow + 1
        
        ' C�񂪋󔒂ɂȂ�����I���i35�s�ڈȊO�j
        If i > 1 And (targetSheet.Cells(currentRow, �ԍϗ\�����).Value = "" Or IsEmpty(targetSheet.Cells(currentRow, �ԍϗ\�����).Value)) Then
            Exit For
        End If
    Next i
    
    �ԍϗ\����擾 = dataArray
End Function

' �o�̓f�[�^�쐬�֐�
' �ԍϗ\�����2���R�[�h�ڂ��烋�[�v���ďo�̓f�[�^���쐬
Public Function �o�̓f�[�^�쐬(targetSheet As Worksheet) As Variant
    Dim �ԍϗ\��f�[�^ As Variant
    Dim ���o���f�[�^ As Variant
    Dim �ؓ������f�[�^ As Variant
    Dim �x�����Q�������f�[�^ As Variant
    Dim �v�Z���ԍŏ��� As Date
    Dim �o�͌���() As Variant
    Dim �o�͍s�� As Long
    Dim i As Long, j As Long, k As Long
    Dim �E�v As String
    
    ' 1. �ԍϗ\����擾
    �ԍϗ\��f�[�^ = �ԍϗ\����擾(targetSheet)
    ���o���f�[�^ = ���o�����S�̎擾(targetSheet)
    �ؓ������f�[�^ = �ؓ������擾(targetSheet)
    �x�����Q�������f�[�^ = �x�����Q�������擾(targetSheet)
    �v�Z���ԍŏ��� = �v�Z���ԍŏ����擾(targetSheet)
    
    ' �f�[�^���݃`�F�b�N
    If Not IsArray(�ԍϗ\��f�[�^) Or UBound(�ԍϗ\��f�[�^, 1) < 2 Then
        Err.Raise 13, "�o�̓f�[�^�쐬", "�ԍϗ\��f�[�^���s�����Ă��܂��B"
    End If
    
    ' �o�͌��ʔz��̏������i�ő�z��s���ŏ������j
    ReDim �o�͌���(1 To 1000, 1 To 19)
    �o�͍s�� = 0
    
    ' 2. �ԍϗ\�����2���R�[�h�ڂ��烋�[�v
    For i = 2 To UBound(�ԍϗ\��f�[�^, 1)
        Dim �ԍϗ\�蓖���f�[�^ As Variant
        Dim �ԍϗ\��O���f�[�^ As Variant
        Dim ���ԊJ�n�� As Date
        Dim ���ԏI���� As Date
        Dim ���������X�g() As Date
        Dim �������� As Long
        
        ' �����ƑO���̃f�[�^��ݒ�
        �ԍϗ\�蓖���f�[�^ = Array(�ԍϗ\��f�[�^(i, 1), �ԍϗ\��f�[�^(i, 2), �ԍϗ\��f�[�^(i, 3))
        �ԍϗ\��O���f�[�^ = Array(�ԍϗ\��f�[�^(i - 1, 1), �ԍϗ\��f�[�^(i - 1, 2), �ԍϗ\��f�[�^(i - 1, 3))
        
        ' 3. �v�Z���Ԃ̍ŏ������擾
        ���ԊJ�n�� = DateSerial(Year(�ԍϗ\�蓖���f�[�^(0)), Month(�ԍϗ\�蓖���f�[�^(0)), 1)
        ���ԊJ�n�� = DateAdd("m", -1, ���ԊJ�n��)
        If ���ԊJ�n�� < �v�Z���ԍŏ��� Then
            ���ԊJ�n�� = �v�Z���ԍŏ���
        End If
        
        ' 4. �v�Z���Ԃ̍ŏI�����擾�i�O1�����̌������j
        ���ԏI���� = DateSerial(Year(�ԍϗ\�蓖���f�[�^(0)), Month(�ԍϗ\�蓖���f�[�^(0)), 1)
        ���ԏI���� = DateAdd("d", -1, ���ԏI����)
        
        ' �������ȍ~�ł���΁A�������̑O���ɐݒ�
        ' �������A���[�������X�e�[�^�X���u�����؂�v�̏ꍇ�͊��������̂��̂�ݒ�
        Dim ������ As Date
        Dim ���[�������X�e�[�^�X As String
        ������ = �������擾(targetSheet)
        ���[�������X�e�[�^�X = ���[�������X�e�[�^�X�擾(targetSheet)
        
        If ���ԏI���� >= ������ Then
            If ���[�������X�e�[�^�X = �����؂ꗝ�R������ Then
                ���ԏI���� = ������
            Else
                ���ԏI���� = DateAdd("d", -1, ������)
            End If
        End If
        
        ' ���������X�g�̏�����
        ReDim ���������X�g(1 To 100)
        �������� = 0
        
        ' 6. �ԍϗ\��O���f�[�^�̓��t�����ԓ��ɂ��邩�`�F�b�N
        Dim �ԍϗ\��O�����t As Date
        �ԍϗ\��O�����t = �ԍϗ\��O���f�[�^(0)
        If �ԍϗ\��O�����t > ���ԊJ�n�� And �ԍϗ\��O�����t <= ���ԏI���� Then
            �������� = �������� + 1
            ���������X�g(��������) = �ԍϗ\��O�����t
        End If
        
        ' 7. ���o�����̓��t�����ԓ��ɂ��邩�`�F�b�N
        If IsArray(���o���f�[�^) And UBound(���o���f�[�^, 1) > 0 Then
            For j = 1 To UBound(���o���f�[�^, 1)
                ' �E�v���u�ԍϕ��v�ŏI���ꍇ�͔��f�ΏۊO�Ƃ���
                �E�v = CStr(���o���f�[�^(j, 2))
                ' �ԍϕ��̏ꍇ�̓X�L�b�v
                If Right(�E�v, Len(�ԍϓE�v�ԍϕ�������)) <> �ԍϓE�v�ԍϕ������� Then
                    Dim ���o���� As Date
                    ���o���� = ���o���f�[�^(j, 1)
                    If ���o���� > ���ԊJ�n�� And ���o���� <= ���ԏI���� Then
                        
                        ' �����̕������Əd�����Ȃ����`�F�b�N
                        Dim �d���t���O As Boolean
                        �d���t���O = False
                        For k = 1 To ��������
                            If ���������X�g(k) = ���o���� Then
                                �d���t���O = True
                                Exit For
                            End If
                        Next k
                        If Not �d���t���O Then
                            �������� = �������� + 1
                            ���������X�g(��������) = ���o����
                        End If
                    End If
                    
                End If
            Next j
        End If
        
        ' 8. �ؓ������f�[�^�̊J�n�������ԓ��ɂ��邩�`�F�b�N
        If IsArray(�ؓ������f�[�^) And UBound(�ؓ������f�[�^, 1) > 0 Then
            For j = 1 To UBound(�ؓ������f�[�^, 1)
                Dim �ؓ������J�n�� As Date
                �ؓ������J�n�� = �ؓ������f�[�^(j, 2)
                If �ؓ������J�n�� > ���ԊJ�n�� And �ؓ������J�n�� <= ���ԏI���� Then
                    ' �����̕������Əd�����Ȃ����`�F�b�N
                    Dim �d���t���O2 As Boolean
                    �d���t���O2 = False
                    For k = 1 To ��������
                        If ���������X�g(k) = �ؓ������J�n�� Then
                            �d���t���O2 = True
                            Exit For
                        End If
                    Next k
                    If Not �d���t���O2 Then
                        �������� = �������� + 1
                        ���������X�g(��������) = �ؓ������J�n��
                    End If
                End If
            Next j
        End If
        
        ' 9. �x�����Q�������f�[�^�̊J�n�������ԓ��ɂ��邩�`�F�b�N
        If IsArray(�x�����Q�������f�[�^) And UBound(�x�����Q�������f�[�^, 1) > 0 Then
            For j = 1 To UBound(�x�����Q�������f�[�^, 1)
                Dim �x�����Q�������J�n�� As Date
                �x�����Q�������J�n�� = �x�����Q�������f�[�^(j, 2)
                If �x�����Q�������J�n�� > ���ԊJ�n�� And �x�����Q�������J�n�� <= ���ԏI���� Then
                    ' �����̕������Əd�����Ȃ����`�F�b�N
                    Dim �d���t���O3 As Boolean
                    �d���t���O3 = False
                    For k = 1 To ��������
                        If ���������X�g(k) = �x�����Q�������J�n�� Then
                            �d���t���O3 = True
                            Exit For
                        End If
                    Next k
                    If Not �d���t���O3 Then
                        �������� = �������� + 1
                        ���������X�g(��������) = �x�����Q�������J�n��
                    End If
                End If
            Next j
        End If
        
        ' ���������\�[�g
        If �������� > 1 Then
            Call �������\�[�g(���������X�g, ��������)
        End If
        
        ' 10. �o�̓��R�[�h�̍쐬
        Dim ���R�[�h�� As Long
        ���R�[�h�� = IIf(�������� = 0, 1, �������� + 1)
        
        For j = 1 To ���R�[�h��
            �o�͍s�� = �o�͍s�� + 1
            
            ' �ʔ�
            �o�͌���(�o�͍s��, �o��_�ʔԗ�) = �o�͍s��
            
            ' �X�e�[�^�X
            �o�͌���(�o�͍s��, �o��_�X�e�[�^�X��) = ����X�e�[�^�X������
            
            ' �C�x���g
            �o�͌���(�o�͍s��, �o��_�C�x���g��) = ���ԍσC�x���g������
            
            ' ���ԍό�
            �o�͌���(�o�͍s��, �o��_���ԍό���) = Format(�ԍϗ\�蓖���f�[�^(0), "yyyy/mm")
            
            ' �v�Z���ԊJ�n��
            If j = 1 Then
                �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = ���ԊJ�n��
            Else
                �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = ���������X�g(j - 1)
            End If
            
            ' �v�Z���ԏI����
            If j = ���R�[�h�� Then
                �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = ���ԏI����
            Else
                �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = DateAdd("d", -1, ���������X�g(j))
            End If
            
            ' ��؂�
            �o�͌���(�o�͍s��, �o��_��؂��) = "�`"
            
            ' �v�Z����
            �o�͌���(�o�͍s��, �o��_�v�Z������) = "=H" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "-F" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "+1"
            
            ' �Ώی����̌v�Z
            Dim �Ώی��� As Double
            Dim �c�� As Double
            Dim ���ؒ����ԍό��� As Double
            
            ' ���o���f�[�^����v�Z���ԊJ�n���Ɠ�������菬�������t�̒��ōő���t�̃f�[�^���擾
            �c�� = 0
            ���ؒ����ԍό��� = 0
            If IsArray(���o���f�[�^) And UBound(���o���f�[�^, 1) > 0 Then
                Dim �v�Z���ԊJ�n��_�Ώی��� As Date
                �v�Z���ԊJ�n��_�Ώی��� = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����)
                
                Dim �ő���t_���o�� As Date
                Dim �ő���t�������� As Boolean
                �ő���t_���o�� = ���t�����l
                �ő���t�������� = False
                
                ' �v�Z���ԊJ�n���Ɠ�������菬�������t�̒��ōő���t��T��
                For k = 1 To UBound(���o���f�[�^, 1)
                    If ���o���f�[�^(k, 1) <= �v�Z���ԊJ�n��_�Ώی��� And ���o���f�[�^(k, 1) > �ő���t_���o�� Then
                        �ő���t_���o�� = ���o���f�[�^(k, 1)
                        �c�� = ���o���f�[�^(k, 4)
                        ���ؒ����ԍό��� = ���o���f�[�^(k, 5)
                        �ő���t�������� = True
                    End If
                Next k
                
                ' �Y������f�[�^��������Ȃ��ꍇ�͍ŏ��̃f�[�^���g�p
                If Not �ő���t�������� And UBound(���o���f�[�^, 1) > 0 Then
                    �c�� = ���o���f�[�^(1, 4)
                    ���ؒ����ԍό��� = ���o���f�[�^(1, 5)
                End If
            End If
            
            �Ώی��� = �c�� - ���ؒ����ԍό���
            
            ' �ԍϗ\���񂩂�v�Z���ԊJ�n���Ɠ�������菬�������t�̒��ōő���t�̃f�[�^�̕ԍό����݌v�����炷
            If IsArray(�ԍϗ\��f�[�^) And UBound(�ԍϗ\��f�[�^, 1) > 0 Then
                Dim �ő���t_�ԍϗ\�� As Date
                Dim �ԍό����݌v_���Z As Double
                �ő���t_�ԍϗ\�� = ���t�����l
                �ԍό����݌v_���Z = 0
                
                For k = 1 To UBound(�ԍϗ\��f�[�^, 1)
                    If �ԍϗ\��f�[�^(k, 1) <= �v�Z���ԊJ�n��_�Ώی��� And �ԍϗ\��f�[�^(k, 1) > �ő���t_�ԍϗ\�� Then
                        �ő���t_�ԍϗ\�� = �ԍϗ\��f�[�^(k, 1)
                        �ԍό����݌v_���Z = �ԍϗ\��f�[�^(k, 3) ' �ԍό����݌v
                    End If
                Next k
                
                �Ώی��� = �Ώی��� - �ԍό����݌v_���Z
            End If
            
            �o�͌���(�o�͍s��, �o��_�Ώی�����) = �Ώی���
            
            ' �����̎擾
            Dim ���� As Double
            Dim ������������ As Boolean
            ���� = 0
            ������������ = False
            
            If IsArray(�ؓ������f�[�^) And UBound(�ؓ������f�[�^, 1) > 0 Then
                Dim �v�Z���ԊJ�n�� As Date
                �v�Z���ԊJ�n�� = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����)
                
                ' �܂��v�Z���ԊJ�n���Ɠ������t�̃f�[�^��T��
                For k = 1 To UBound(�ؓ������f�[�^, 1)
                    If �ؓ������f�[�^(k, 2) = �v�Z���ԊJ�n�� Then
                        ���� = �ؓ������f�[�^(k, 1)
                        ������������ = True
                        Exit For
                    End If
                Next k
                
                ' �������t�̃f�[�^���Ȃ��ꍇ�A�v�Z���ԊJ�n����菬�������t�̒��ōł��傫�����t��T���i����͔C�ӂ̓��t���󂯓���j
                If Not ������������ Then
                    Dim �ő���t As Date
                    �ő���t = ���t�����l ' �����l�Ƃ��čŏ����t��ݒ�
                    
                    For k = 1 To UBound(�ؓ������f�[�^, 1)
                        If �ؓ������f�[�^(k, 2) < �v�Z���ԊJ�n�� And (�ؓ������f�[�^(k, 2) > �ő���t Or (�ؓ������f�[�^(k, 2) = �ő���t And �ő���t = ���t�����l)) Then
                            �ő���t = �ؓ������f�[�^(k, 2)
                            ���� = �ؓ������f�[�^(k, 1)
                            ������������ = True
                        End If
                    Next k
                End If
            End If
            �o�͌���(�o�͍s��, �o��_������) = ����
            
            ' �ϐ��̐����ݒ�i�Ώی����~�����~�v�Z�����j
            �o�͌���(�o�͍s��, �o��_�ϐ���) = "=E" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*J" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*I" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g)
            
            ' �������z�̐����ݒ�
            Dim ���ݍs�ԍ� As Long
            ���ݍs�ԍ� = �o�͍s�� + �o�͊J�n�s�I�t�Z�b�g
            
            If j = 1 Then
                ' J=1�̏ꍇ�F=ROUNDDOWN(K�s�ԍ�/365,0)
                �o�͌���(�o�͍s��, �o��_�������z��) = "=ROUNDDOWN(K" & ���ݍs�ԍ� & "/365,0)"
            Else
                ' J=1�ȊO�̏ꍇ�F=ROUNDDOWN(SUM(K(J=1���̍s�ԍ�):K���݂̍s�ԍ�)/365,0)-SUM(L(J=1���̍s�ԍ�):L���݂̍s�ԍ�-1)
                Dim J1�J�n�s�ԍ� As Long
                J1�J�n�s�ԍ� = (�o�͍s�� - j + 1) + �o�͊J�n�s�I�t�Z�b�g ' J=1���̍s�ԍ����v�Z
                �o�͌���(�o�͍s��, �o��_�������z��) = "=ROUNDDOWN(SUM(K" & J1�J�n�s�ԍ� & ":K" & ���ݍs�ԍ� & ")/365,0)-SUM(L" & J1�J�n�s�ԍ� & ":L" & (���ݍs�ԍ� - 1) & ")"
            End If
            
            ' �x�����Q���͋�
            �o�͌���(�o�͍s��, �o��_�x�����Q����) = ""
            
            ' �ؓ��E�ԍϏ��̐ݒ�
            Dim ���o�����S�̃f�[�^_�ؓ��ԍ� As Variant
            ���o�����S�̃f�[�^_�ؓ��ԍ� = ���o�����S�̎擾(targetSheet)
            
            If IsArray(���o�����S�̃f�[�^_�ؓ��ԍ�) And UBound(���o�����S�̃f�[�^_�ؓ��ԍ�, 1) > 0 Then
                Dim �ؓ��ݒ�ς� As Boolean
                Dim �ԍϐݒ�ς� As Boolean
                �ؓ��ݒ�ς� = False
                �ԍϐݒ�ς� = False
                
                For k = 1 To UBound(���o�����S�̃f�[�^_�ؓ��ԍ�, 1)
                    ' �v�Z���ԊJ�n���ƈ�v����ꍇ�̂ݏ���
                    If ���o�����S�̃f�[�^_�ؓ��ԍ�(k, 1) = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) Then
                        �E�v = CStr(���o�����S�̃f�[�^_�ؓ��ԍ�(k, 2))
                        
                        ' �ؓ����̐ݒ�i�ؓ��܂��͎؊��j
                        If Not �ؓ��ݒ�ς� And ((Len(�E�v) >= Len(�ؓ��E�v�ؓ�������) And Right(�E�v, Len(�ؓ��E�v�ؓ�������)) = �ؓ��E�v�ؓ�������) Or (Len(�E�v) >= Len(�ؓ��E�v�؊�������) And Right(�E�v, Len(�ؓ��E�v�؊�������)) = �ؓ��E�v�؊�������)) Then
                            �o�͌���(�o�͍s��, �o��_�ؓ�����) = ���o�����S�̃f�[�^_�ؓ��ԍ�(k, 1)
                            �o�͌���(�o�͍s��, �o��_�ؓ��z��) = ���o�����S�̃f�[�^_�ؓ��ԍ�(k, 3)
                            �ؓ��ݒ�ς� = True
                        End If
                        
                        ' �ԍϏ��̐ݒ�
                        If Not �ԍϐݒ�ς� And Len(�E�v) >= Len(�ԍϓE�v�ԍϕ�������) And Right(�E�v, Len(�ԍϓE�v�ԍϕ�������)) = �ԍϓE�v�ԍϕ������� Then
                            �o�͌���(�o�͍s��, �o��_�ԍϓ���) = ���o�����S�̃f�[�^_�ؓ��ԍ�(k, 1)
                            �o�͌���(�o�͍s��, �o��_����_�ԍϊz��) = ���o�����S�̃f�[�^_�ؓ��ԍ�(k, 3)
                            �ԍϐݒ�ς� = True
                        End If
                        
                        ' �����ݒ�ς݂̏ꍇ�̓��[�v���I��
                        If �ؓ��ݒ�ς� And �ԍϐݒ�ς� Then
                            Exit For
                        End If
                    End If
                Next k
            End If

            
        Next j
        
        ' 11. ���ؕ����R�[�h�̒l��ݒ肷��
        Dim ������_���� As Date
        ������_���� = �������擾(targetSheet)
        
        ' �ԍϗ\��f�[�^�̓��t����������菬�����ꍇ�̂݉��ؕ����R�[�h���쐬
        If �ԍϗ\�蓖���f�[�^(0) < ������_���� Then
            ' ���ؕ����R�[�h�̊��Ԑݒ�
            Dim ���؊��ԊJ�n�� As Date
            Dim ���؊��ԏI���� As Date
            ���؊��ԊJ�n�� = �ԍϗ\�蓖���f�[�^(0)
            
            ' �v�Z���ԏI�����i�������̑O���A���������[�������X�e�[�^�X���u�����؂�v�̏ꍇ�͊��������̂��́j
            If ���[�������X�e�[�^�X = �����؂ꗝ�R������ Then
                ���؊��ԏI���� = ������_����
            Else
                ���؊��ԏI���� = DateAdd("d", -1, ������_����)
            End If
            
            ' ���ؕ����R�[�h�p�̕��������X�g���쐬
            Dim ���ؕ��������X�g(1 To 100) As Date
            Dim ���ؕ������� As Long
            ���ؕ������� = 0
            
            ' �x�����Q�������f�[�^�̊J�n�������ԓ��ɂ��邩�`�F�b�N
            If IsArray(�x�����Q�������f�[�^) And UBound(�x�����Q�������f�[�^, 1) > 0 Then
                For j = 1 To UBound(�x�����Q�������f�[�^, 1)
                    Dim �x�����Q�������J�n��_���� As Date
                    �x�����Q�������J�n��_���� = �x�����Q�������f�[�^(j, 2)
                    If �x�����Q�������J�n��_���� > ���؊��ԊJ�n�� And �x�����Q�������J�n��_���� <= ���؊��ԏI���� Then
                        ' �����̕������Əd�����Ȃ����`�F�b�N
                        Dim �d���t���O_���� As Boolean
                        �d���t���O_���� = False
                        For k = 1 To ���ؕ�������
                            If ���ؕ��������X�g(k) = �x�����Q�������J�n��_���� Then
                                �d���t���O_���� = True
                                Exit For
                            End If
                        Next k
                        If Not �d���t���O_���� Then
                            ���ؕ������� = ���ؕ������� + 1
                            ���ؕ��������X�g(���ؕ�������) = �x�����Q�������J�n��_����
                        End If
                    End If
                Next j
            End If
            
            ' ���������X�g���\�[�g
            If ���ؕ������� > 1 Then
                For j = 1 To ���ؕ������� - 1
                    For k = j + 1 To ���ؕ�������
                        If ���ؕ��������X�g(j) > ���ؕ��������X�g(k) Then
                            Dim temp_���� As Date
                            temp_���� = ���ؕ��������X�g(j)
                            ���ؕ��������X�g(j) = ���ؕ��������X�g(k)
                            ���ؕ��������X�g(k) = temp_����
                        End If
                    Next k
                Next j
            End If
            
            ' ���ؕ����R�[�h�𕪊����č쐬
            Dim ���؃Z�O�����g�J�n�� As Date
            Dim ���؃Z�O�����g�I���� As Date
            ���؃Z�O�����g�J�n�� = ���؊��ԊJ�n��
            
            For j = 0 To ���ؕ�������
                ' �Z�O�����g�I�����̐ݒ�
                If j = ���ؕ������� Then
                    ���؃Z�O�����g�I���� = ���؊��ԏI����
                Else
                    ���؃Z�O�����g�I���� = DateAdd("d", -1, ���ؕ��������X�g(j + 1))
                End If
                
                ' �Z�O�����g���L���ȏꍇ�̂݃��R�[�h���쐬
                If ���؃Z�O�����g�J�n�� <= ���؃Z�O�����g�I���� Then
                    �o�͍s�� = �o�͍s�� + 1
                    
                    ' �ʔ�
                    �o�͌���(�o�͍s��, �o��_�ʔԗ�) = �o�͍s��
                    
                    ' �X�e�[�^�X
                    �o�͌���(�o�͍s��, �o��_�X�e�[�^�X��) = ���؃C�x���g������
                    
                    ' �C�x���g
                    �o�͌���(�o�͍s��, �o��_�C�x���g��) = ���ԍσC�x���g������
                    
                    ' ���ԍό�
                    �o�͌���(�o�͍s��, �o��_���ԍό���) = Format(�ԍϗ\�蓖���f�[�^(0), "yyyy/mm")
                    
                    ' �v�Z���ԊJ�n��
                    �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = ���؃Z�O�����g�J�n��
                    
                    ' �v�Z���ԏI����
                    �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = ���؃Z�O�����g�I����
                    
                    ' ��؂�
                    �o�͌���(�o�͍s��, �o��_��؂��) = "�`"
                    
                    ' �v�Z����
                    �o�͌���(�o�͍s��, �o��_�v�Z������) = "=H" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "-F" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "+1"
                    
                    ' �Ώی���
                    �o�͌���(�o�͍s��, �o��_�Ώی�����) = �ԍϗ\�蓖���f�[�^(1) ' �ԍό���
                    
                    ' �x�����Q�������̎擾
                    Dim �x�����Q������_���� As Double
                    Dim �x�����Q��������������_���� As Boolean
                    �x�����Q������_���� = 0
                    �x�����Q��������������_���� = False
                    
                    If IsArray(�x�����Q�������f�[�^) And UBound(�x�����Q�������f�[�^, 1) > 0 Then
                        ' �܂��v�Z���ԊJ�n���Ɠ������t�̃f�[�^��T��
                        For k = 1 To UBound(�x�����Q�������f�[�^, 1)
                            If �x�����Q�������f�[�^(k, 2) = ���؃Z�O�����g�J�n�� Then
                                �x�����Q������_���� = �x�����Q�������f�[�^(k, 1)
                                �x�����Q��������������_���� = True
                                Exit For
                            End If
                        Next k
                        
                        ' �������t�̃f�[�^���Ȃ��ꍇ�A�v�Z���ԊJ�n����菬�������t�̒��ōł��傫�����t��T��
                        If Not �x�����Q��������������_���� Then
                            Dim �ő���t_�x�����Q��_���� As Date
                            �ő���t_�x�����Q��_���� = ���t�����l
                            
                            For k = 1 To UBound(�x�����Q�������f�[�^, 1)
                                If �x�����Q�������f�[�^(k, 2) < ���؃Z�O�����g�J�n�� And (�x�����Q�������f�[�^(k, 2) > �ő���t_�x�����Q��_���� Or (�x�����Q�������f�[�^(k, 2) = �ő���t_�x�����Q��_���� And �ő���t_�x�����Q��_���� = ���t�����l)) Then
                                    �ő���t_�x�����Q��_���� = �x�����Q�������f�[�^(k, 2)
                                    �x�����Q������_���� = �x�����Q�������f�[�^(k, 1)
                                    �x�����Q��������������_���� = True
                                End If
                            Next k
                        End If
                    End If
                    �o�͌���(�o�͍s��, �o��_������) = �x�����Q������_����
                    
                    ' �ϐ��̐����ݒ�i�Ώی����~�����~�v�Z�����j
                    �o�͌���(�o�͍s��, �o��_�ϐ���) = "=E" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*J" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*I" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g)
                    
                    ' �������z�͋�
                    �o�͌���(�o�͍s��, �o��_�������z��) = ""
                    
                    ' �x�����Q���̐����ݒ�i�ϐ�/365�j
                    �o�͌���(�o�͍s��, �o��_�x�����Q����) = "=ROUNDDOWN(K" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "/365,0)"
                    
                    ' �ؓ��E�ԍϏ��̐ݒ�i�ŏ��̃Z�O�����g�̂݁j
                    If j = 0 Then
                        Dim ���o�����S�̃f�[�^_�ؓ��ԍ�_���� As Variant
                        ���o�����S�̃f�[�^_�ؓ��ԍ�_���� = ���o�����S�̎擾(targetSheet)
                        
                        If IsArray(���o�����S�̃f�[�^_�ؓ��ԍ�_����) And UBound(���o�����S�̃f�[�^_�ؓ��ԍ�_����, 1) > 0 Then
                            Dim k_���� As Long
                            Dim �ؓ��ݒ�ς�_���� As Boolean
                            Dim �ԍϐݒ�ς�_���� As Boolean
                            �ؓ��ݒ�ς�_���� = False
                            �ԍϐݒ�ς�_���� = False
                            
                            For k_���� = 1 To UBound(���o�����S�̃f�[�^_�ؓ��ԍ�_����, 1)
                                ' �v�Z���ԊJ�n���ƈ�v����ꍇ�̂ݏ���
                                If ���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 1) = ���؃Z�O�����g�J�n�� Then
                                    Dim �E�v_���� As String
                                    �E�v_���� = CStr(���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 2))
                                    
                                    ' �ؓ����̐ݒ�i�ؓ��܂��͎؊��j
                                    If Not �ؓ��ݒ�ς�_���� And ((Len(�E�v_����) >= Len(�ؓ��E�v�ؓ�������) And Right(�E�v_����, Len(�ؓ��E�v�ؓ�������)) = �ؓ��E�v�ؓ�������) Or (Len(�E�v_����) >= Len(�ؓ��E�v�؊�������) And Right(�E�v_����, Len(�ؓ��E�v�؊�������)) = �ؓ��E�v�؊�������)) Then
                                        �o�͌���(�o�͍s��, �o��_�ؓ�����) = ���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 1)
                                        �o�͌���(�o�͍s��, �o��_�ؓ��z��) = ���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 3)
                                        �ؓ��ݒ�ς�_���� = True
                                    End If
                                    
                                    ' �ԍϏ��̐ݒ�
                                    If Not �ԍϐݒ�ς�_���� And Len(�E�v_����) >= Len(�ԍϓE�v�ԍϕ�������) And Right(�E�v_����, Len(�ԍϓE�v�ԍϕ�������)) = �ԍϓE�v�ԍϕ������� Then
                                        �o�͌���(�o�͍s��, �o��_�ԍϓ���) = ���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 1)
                                        �o�͌���(�o�͍s��, �o��_����_�ԍϊz��) = ���o�����S�̃f�[�^_�ؓ��ԍ�_����(k_����, 3)
                                        �ԍϐݒ�ς�_���� = True
                                    End If
                                    
                                    ' �����ݒ�ς݂̏ꍇ�̓��[�v���I��
                                    If �ؓ��ݒ�ς�_���� And �ԍϐݒ�ς�_���� Then
                                        Exit For
                                    End If
                                End If
                            Next k_����
                        End If
                    End If
                    
                    ' ���̃Z�O�����g�̊J�n����ݒ�
                    If j < ���ؕ������� Then
                        ���؃Z�O�����g�J�n�� = ���ؕ��������X�g(j + 1)
                    End If
                End If
            Next j
        End If
        
    Next i

    ' 12. �������R�[�h�̏���
    Dim ������_�������R�[�h As Date
    Dim ��� As Date
    Dim �葱�J�n��_�������R�[�h As Date
    Dim ���o�����S�̃f�[�^ As Variant
    Dim �ԍϗ����f�[�^ As Variant
    
    ������_�������R�[�h = �������擾(targetSheet)
    ��� = Date - 1
    �葱�J�n��_�������R�[�h = �葱�J�n���擾(targetSheet)
    ���o�����S�̃f�[�^ = ���o�����S�̎擾(targetSheet)
    �ԍϗ����f�[�^ = �ԍϗ������擾(targetSheet)
    
    ' ���������X�g�̍쐬
    Dim �������������X�g() As Date
    Dim ������������ As Long
    ReDim �������������X�g(1 To 100)
    ������������ = 0
    
    ' �葱�J�n�������ԓ��ɂ��邩�`�F�b�N
    If �葱�J�n��_�������R�[�h >= ������_�������R�[�h And �葱�J�n��_�������R�[�h <= ��� Then
        ������������ = ������������ + 1
        �������������X�g(������������) = �葱�J�n��_�������R�[�h
    End If
    
    ' ���o�����S�̎擾�̓��t�����ԓ��ɂ��邩�`�F�b�N
    If IsArray(���o�����S�̃f�[�^) And UBound(���o�����S�̃f�[�^, 1) > 0 Then
        For j = 1 To UBound(���o�����S�̃f�[�^, 1)
            Dim ���o����_���� As Date
            ���o����_���� = ���o�����S�̃f�[�^(j, 1)
            If ���o����_���� >= ������_�������R�[�h And ���o����_���� <= ��� Then
                ' �����̕������Əd�����Ȃ����`�F�b�N
                Dim �d���t���O_���� As Boolean
                �d���t���O_���� = False
                For k = 1 To ������������
                    If �������������X�g(k) = ���o����_���� Then
                        �d���t���O_���� = True
                        Exit For
                    End If
                Next k
                If Not �d���t���O_���� Then
                    ������������ = ������������ + 1
                    �������������X�g(������������) = ���o����_����
                End If
            End If
        Next j
    End If
    
    ' �ԍϗ����̓��t�����ԓ��ɂ��邩�`�F�b�N
    If IsArray(�ԍϗ����f�[�^) And UBound(�ԍϗ����f�[�^, 1) > 0 Then
        For j = 1 To UBound(�ԍϗ����f�[�^, 1)
            Dim �ԍϗ����_���� As Date
            �ԍϗ����_���� = �ԍϗ����f�[�^(j, 1)
            If �ԍϗ����_���� >= ������_�������R�[�h And �ԍϗ����_���� <= ��� Then
                ' �����̕������Əd�����Ȃ����`�F�b�N
                Dim �d���t���O2_���� As Boolean
                �d���t���O2_���� = False
                For k = 1 To ������������
                    If �������������X�g(k) = �ԍϗ����_���� Then
                        �d���t���O2_���� = True
                        Exit For
                    End If
                Next k
                If Not �d���t���O2_���� Then
                    ������������ = ������������ + 1
                    �������������X�g(������������) = �ԍϗ����_����
                End If
            End If
        Next j
    End If
    
    ' ���������\�[�g
    If ������������ > 1 Then
        Call �������\�[�g(�������������X�g, ������������)
    End If
    
    ' �������R�[�h�̍쐬
    Dim �������R�[�h�� As Long
    �������R�[�h�� = IIf(������������ = 0, 1, ������������ + 1)
    
    ' J=1���̍s�ԍ����L�^
    Dim J1���̍s�ԍ�_���� As Long
    J1���̍s�ԍ�_���� = �o�͍s�� + 1
    
    For j = 1 To �������R�[�h��
        �o�͍s�� = �o�͍s�� + 1
        
        ' �ʔ�
        �o�͌���(�o�͍s��, �o��_�ʔԗ�) = �o�͍s��
        
        ' �v�Z���ԊJ�n��
        If j = 1 Then
            �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = ������_�������R�[�h
        Else
            �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = �������������X�g(j - 1)
        End If
        
        ' �v�Z���ԏI����
        If j = �������R�[�h�� Then
            �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = ���
        Else
            �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = DateAdd("d", -1, �������������X�g(j))
        End If
        
        ' �X�e�[�^�X
        If �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = �葱�J�n��_�������R�[�h Then
            �o�͌���(�o�͍s��, �o��_�X�e�[�^�X��) = �������X�e�[�^�X������
        Else
            �o�͌���(�o�͍s��, �o��_�X�e�[�^�X��) = �����X�e�[�^�X������
        End If
        
        ' �C�x���g
        If �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = �葱�J�n��_�������R�[�h Then
            �o�͌���(�o�͍s��, �o��_�C�x���g��) = �j�Y�C�x���g������
        ElseIf �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) = ������_�������R�[�h Then
            �o�͌���(�o�͍s��, �o��_�C�x���g��) = �������R�擾(targetSheet)
        Else
            �o�͌���(�o�͍s��, �o��_�C�x���g��) = �����C�x���g������
        End If
        
        ' ���ԍό�
        If �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������) = ��� Then
            �o�͌���(�o�͍s��, �o��_���ԍό���) = "�v�Z��"
        Else
            �o�͌���(�o�͍s��, �o��_���ԍό���) = "�["
        End If
        
        ' �Ώی���
        Dim �Ώی���_���� As Double
        �Ώی���_���� = 0
        If IsArray(���o�����S�̃f�[�^) And UBound(���o�����S�̃f�[�^, 1) > 0 Then
            Dim �v�Z���ԊJ�n��_���� As Date
            �v�Z���ԊJ�n��_���� = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����)
            
            ' �������t�̃f�[�^��T��
            Dim �������t��������_���� As Boolean
            �������t��������_���� = False
            For k = 1 To UBound(���o�����S�̃f�[�^, 1)
                If ���o�����S�̃f�[�^(k, 1) = �v�Z���ԊJ�n��_���� Then
                    �Ώی���_���� = ���o�����S�̃f�[�^(k, 4) ' �c��
                    �������t��������_���� = True
                    Exit For
                End If
            Next k
            
            ' �������t���Ȃ��ꍇ�A�v�Z���ԊJ�n����菬�������t�̒��ōő�̂��̂�T��
            If Not �������t��������_���� Then
                Dim �ő���t_���� As Date
                �ő���t_���� = ���t�����l
                For k = 1 To UBound(���o�����S�̃f�[�^, 1)
                    If ���o�����S�̃f�[�^(k, 1) < �v�Z���ԊJ�n��_���� And ���o�����S�̃f�[�^(k, 1) > �ő���t_���� Then
                        �ő���t_���� = ���o�����S�̃f�[�^(k, 1)
                        �Ώی���_���� = ���o�����S�̃f�[�^(k, 4) ' �c��
                    End If
                Next k
            End If
        End If
        �o�͌���(�o�͍s��, �o��_�Ώی�����) = �Ώی���_����
        
        ' �v�Z����
        �o�͌���(�o�͍s��, �o��_�v�Z������) = "=H" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "-F" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "+1"
        
        ' ��؂�
        �o�͌���(�o�͍s��, �o��_��؂��) = "�`"
        
        ' �����i�x�����Q�������j
        Dim �x�����Q������_���� As Double
        �x�����Q������_���� = 0
        If IsArray(�x�����Q�������f�[�^) And UBound(�x�����Q�������f�[�^, 1) > 0 Then
            Dim �v�Z���ԊJ�n��_����_���� As Date
            �v�Z���ԊJ�n��_����_���� = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����)
            
            ' �������t�̃f�[�^��T��
            Dim ������������_���� As Boolean
            ������������_���� = False
            For k = 1 To UBound(�x�����Q�������f�[�^, 1)
                If �x�����Q�������f�[�^(k, 2) = �v�Z���ԊJ�n��_����_���� Then
                    �x�����Q������_���� = �x�����Q�������f�[�^(k, 1)
                    ������������_���� = True
                    Exit For
                End If
            Next k
            
            ' �������t���Ȃ��ꍇ�A�v�Z���ԊJ�n����菬�������t�̒��ōł��傫�����t��T���i����͔C�ӂ̓��t���󂯓���j
            If Not ������������_���� Then
                Dim �ő���t_����_���� As Date
                �ő���t_����_���� = ���t�����l ' �����l�Ƃ��čŏ����t��ݒ�
                
                For k = 1 To UBound(�x�����Q�������f�[�^, 1)
                    If �x�����Q�������f�[�^(k, 2) < �v�Z���ԊJ�n��_����_���� And (�x�����Q�������f�[�^(k, 2) > �ő���t_����_���� Or (�x�����Q�������f�[�^(k, 2) = �ő���t_����_���� And �ő���t_����_���� = ���t�����l)) Then
                        �ő���t_����_���� = �x�����Q�������f�[�^(k, 2)
                        �x�����Q������_���� = �x�����Q�������f�[�^(k, 1)
                    End If
                Next k
            End If
        End If
        �o�͌���(�o�͍s��, �o��_������) = �x�����Q������_����
        
        ' �ϐ��̐����ݒ�
        �o�͌���(�o�͍s��, �o��_�ϐ���) = "=E" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*J" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "*I" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g)
        
        ' �������z�͋�
        �o�͌���(�o�͍s��, �o��_�������z��) = ""
        
        ' �x�����Q���̐����ݒ�
        If j = 1 Then
            ' J=1�̏ꍇ�F=ROUNDDOWN(K�s�ԍ�/365,0)
            �o�͌���(�o�͍s��, �o��_�x�����Q����) = "=ROUNDDOWN(K" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & "/365,0)"
        Else
            ' J=1�ȊO�̏ꍇ�F=ROUNDDOWN(SUM(K(J=1���̍s�ԍ�):K���݂̍s�ԍ�)/365,0)-SUM(L(J=1���̍s�ԍ�):M���݂̍s�ԍ�-1)
            �o�͌���(�o�͍s��, �o��_�x�����Q����) = "=ROUNDDOWN(SUM(K" & (J1���̍s�ԍ�_���� + �o�͊J�n�s�I�t�Z�b�g) & ":K" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g) & ")/365,0)-SUM(M" & (J1���̍s�ԍ�_���� + �o�͊J�n�s�I�t�Z�b�g) & ":M" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g - 1) & ")"
        End If
        
        ' �ԍϓ�
        Dim �ԍϓ�_���� As Variant
        �ԍϓ�_���� = ""
        Dim �v�Z���ԊJ�n��_�ԍϓ� As Date
        �v�Z���ԊJ�n��_�ԍϓ� = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����)
        
        ' ���o�����S�̎擾�œ������t�����邩�`�F�b�N
        If IsArray(���o�����S�̃f�[�^) And UBound(���o�����S�̃f�[�^, 1) > 0 Then
            For k = 1 To UBound(���o�����S�̃f�[�^, 1)
                If ���o�����S�̃f�[�^(k, 1) = �v�Z���ԊJ�n��_�ԍϓ� Then
                    �ԍϓ�_���� = �v�Z���ԊJ�n��_�ԍϓ�
                    Exit For
                End If
            Next k
        End If
        
        ' �ԍϗ����œ������t�����邩�`�F�b�N
        If �ԍϓ�_���� = "" And IsArray(�ԍϗ����f�[�^) And UBound(�ԍϗ����f�[�^, 1) > 0 Then
            For k = 1 To UBound(�ԍϗ����f�[�^, 1)
                If �ԍϗ����f�[�^(k, 1) = �v�Z���ԊJ�n��_�ԍϓ� Then
                    �ԍϓ�_���� = �v�Z���ԊJ�n��_�ԍϓ�
                    Exit For
                End If
            Next k
        End If
        �o�͌���(�o�͍s��, �o��_�ԍϓ���) = �ԍϓ�_����
        
        ' ����_�ԍϊz
        Dim �����ԍϊz_���� As Variant
        �����ԍϊz_���� = ""
        If IsArray(���o�����S�̃f�[�^) And UBound(���o�����S�̃f�[�^, 1) > 0 Then
            For k = 1 To UBound(���o�����S�̃f�[�^, 1)
                If ���o�����S�̃f�[�^(k, 1) = �v�Z���ԊJ�n��_�ԍϓ� Then
                    �����ԍϊz_���� = ���o�����S�̃f�[�^(k, 3) ' ���o�����z
                    Exit For
                End If
            Next k
        End If
        �o�͌���(�o�͍s��, �o��_����_�ԍϊz��) = �����ԍϊz_����
        
        ' ����_�ԍϊz
        Dim �����ԍϊz_���� As Variant
        �����ԍϊz_���� = ""
        If IsArray(�ԍϗ����f�[�^) And UBound(�ԍϗ����f�[�^, 1) > 0 Then
            For k = 1 To UBound(�ԍϗ����f�[�^, 1)
                If �ԍϗ����f�[�^(k, 1) = �v�Z���ԊJ�n��_�ԍϓ� And InStr(�ԍϗ����f�[�^(k, 2), �����E�v������) > 0 Then
                    �����ԍϊz_���� = �ԍϗ����f�[�^(k, 3) ' �o�����z
                    Exit For
                End If
            Next k
        End If
        �o�͌���(�o�͍s��, �o��_����_�ԍϊz��) = �����ԍϊz_����
        
        ' �x����_�ԍϊz
        Dim �x�����ԍϊz_���� As Variant
        �x�����ԍϊz_���� = ""
        If IsArray(�ԍϗ����f�[�^) And UBound(�ԍϗ����f�[�^, 1) > 0 Then
            For k = 1 To UBound(�ԍϗ����f�[�^, 1)
                If �ԍϗ����f�[�^(k, 1) = �v�Z���ԊJ�n��_�ԍϓ� And InStr(�ԍϗ����f�[�^(k, 2), �x�����Q���E�v������) > 0 Then
                    �x�����ԍϊz_���� = �ԍϗ����f�[�^(k, 3) ' �o�����z
                    Exit For
                End If
            Next k
        End If
        �o�͌���(�o�͍s��, �o��_�x����_�ԍϊz��) = �x�����ԍϊz_����
        
    Next j
    
    ' �������R�[�h�́u�v�v�s��ǉ�
    �o�͍s�� = �o�͍s�� + 1
    
    ' �ʔԁi�ݒ肵�Ȃ��j
    ' �o�͌���(�o�͍s��, �o��_�ʔԗ�) = �o�͍s��
    
    ' ���ԍό�
    �o�͌���(�o�͍s��, �o��_���ԍό���) = "�v"
    
    ' �Ώی����i�O���R�[�h�̒l�j
    If �o�͍s�� > 1 Then
        �o�͌���(�o�͍s��, �o��_�Ώی�����) = �o�͌���(�o�͍s�� - 1, �o��_�Ώی�����)
    End If
    
    ' �������z�̐����ݒ�i���R�[�h�S�̗̂������z�̍��v-���R�[�h�S�̗̂���_�ԍϊz�̍��v�j
    �o�͌���(�o�͍s��, �o��_�������z��) = "=SUM(L" & (1 + �o�͊J�n�s�I�t�Z�b�g) & ":L" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g - 1) & ")-SUM(R" & (1 + �o�͊J�n�s�I�t�Z�b�g) & ":R" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g - 1) & ")"
    
    ' �x�����Q���̐����ݒ�i���R�[�h�S�̂̒x�����Q���̍��v-���R�[�h�S�̂̒x����_�ԍϊz�̍��v�j
    �o�͌���(�o�͍s��, �o��_�x�����Q����) = "=SUM(M" & (1 + �o�͊J�n�s�I�t�Z�b�g) & ":M" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g - 1) & ")-SUM(S" & (1 + �o�͊J�n�s�I�t�Z�b�g) & ":S" & (�o�͍s�� + �o�͊J�n�s�I�t�Z�b�g - 1) & ")"
    
    ' ���ʔz��̃T�C�Y�𒲐�
    If �o�͍s�� > 0 Then
        ' �V�����z����쐬���ĕK�v�ȕ������R�s�[
        Dim �ŏI����() As Variant
        ReDim �ŏI����(1 To �o�͍s��, 1 To 19)
        
        Dim copyRow As Long, copyCol As Long
        For copyRow = 1 To �o�͍s��
            For copyCol = 1 To 19
                �ŏI����(copyRow, copyCol) = �o�͌���(copyRow, copyCol)
            Next copyCol
        Next copyRow
        
        �o�̓f�[�^�쐬 = �ŏI����
    Else
        �o�̓f�[�^�쐬 = Array()
    End If
End Function

' ���������\�[�g����w���p�[�֐�
Private Sub �������\�[�g(���������X�g() As Date, �������� As Long)
    Dim i As Long, j As Long
    Dim temp As Date
    
    For i = 1 To �������� - 1
        For j = i + 1 To ��������
            If ���������X�g(i) > ���������X�g(j) Then
                temp = ���������X�g(i)
                ���������X�g(i) = ���������X�g(j)
                ���������X�g(j) = temp
            End If
        Next j
    Next i
End Sub



' ���o�����S�̎擾�֐��i�ԍϕ����܂ށj
' �w�肳�ꂽ�V�[�g�̓��o���J�n�s����S�Ă̓��o�������擾�i�ԍϕ��̏��O�Ȃ��j
Public Function ���o�����S�̎擾(targetSheet As Worksheet) As Variant
    Dim startRow As Long
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long, j As Long
    
    startRow = ���o���J�n�s ' �J�n�s
    currentRow = startRow
    rowCount = 0
    
    ' �f�[�^�s�����J�E���g�iB�񂪋󔒂ɂȂ�܂ŁA�S�Ă̍s���J�E���g�j
    Do While targetSheet.Cells(currentRow, ���o������).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
    
    ' �f�[�^�����݂��Ȃ��ꍇ�͋�̔z���Ԃ�
    If rowCount = 0 Then
        ���o�����S�̎擾 = Array()
        Exit Function
    End If
    
    ' �z����������i�s�� x 5��j
    ReDim dataArray(1 To rowCount, 1 To 5)
    
    ' �f�[�^���擾���ăo���f�[�V����
    currentRow = startRow
    
    For i = 1 To rowCount
        ' B��F���o�����i���t�`�F�b�N�j
        Dim dateValue As Variant
        dateValue = targetSheet.Cells(currentRow, ���o������).Value
        If Not IsDate(dateValue) Then
            Err.Raise 13, "���o�����S�̎擾", currentRow & "�s�ڂ�B��i���o�����j�����t�ł͂���܂���B"
        End If
        dataArray(i, 1) = CDate(dateValue)
        
        ' C��F�E�v�i������A�`�F�b�N�s�v�j
        dataArray(i, 2) = CStr(targetSheet.Cells(currentRow, �E�v��).Value)
        
        ' ������t�f�[�^�����F���݂��ԍϕ��œ�����t�̔�ԍϕ��f�[�^�����݂���ꍇ�̓G���[
        Dim currentDate As Date
        currentDate = dataArray(i, 1)
        Dim currentRemark As String
        currentRemark = dataArray(i, 2)
        Dim isCurrentRepayment As Boolean
        isCurrentRepayment = (Len(currentRemark) >= Len(�ԍϓE�v�ԍϕ�������) And Right(currentRemark, Len(�ԍϓE�v�ԍϕ�������)) = �ԍϓE�v�ԍϕ�������)
        
        ' �ȑO�̃f�[�^�ɓ�����t�����邩�`�F�b�N
        dim checkIndex as long
        For checkIndex = 1 To i - 1
            If dataArray(checkIndex, 1) = currentDate Then
                Dim previousRemark As String
                previousRemark = CStr(dataArray(checkIndex, 2))
                Dim isPreviousRepayment As Boolean
                isPreviousRepayment = (Len(previousRemark) >= Len(�ԍϓE�v�ԍϕ�������) And Right(previousRemark, Len(�ԍϓE�v�ԍϕ�������)) = �ԍϓE�v�ԍϕ�������)
                
                ' �����Ƃ��ԍϕ��łȂ��ꍇ�̓G���[
                If Not isCurrentRepayment And Not isPreviousRepayment Then
                    Err.Raise 13, "���o�����S�̎擾", "������t�i" & Format(currentDate, "yyyy/mm/dd") & "�j�ɕ����̔�ԍϕ��f�[�^�����݂��܂��B�f�[�^���m�F���Ă��������B"
                End If
                
                ' ���݂��ԍϕ��ňȑO����ԍϕ��̏ꍇ�A���݂̃��R�[�h���X�L�b�v
                If isCurrentRepayment And Not isPreviousRepayment Then
                    GoTo NextRecord
                End If
                
                ' �ȑO���ԍϕ��Ō��݂���ԍϕ��̏ꍇ�A�ȑO�̃��R�[�h���폜�i�����Ƃ��ă}�[�N�j
                If Not isCurrentRepayment And isPreviousRepayment Then
                    ' �ȑO�̃��R�[�h�𖳌��Ƃ��ă}�[�N�i���t���ŏ��l�ɐݒ�j
                    dataArray(checkIndex, 1) = ���t�����l
                End If
            End If
        Next checkIndex
        
        ' D��F���o�����z�i���l�`�F�b�N�j
        Dim amountValue As Variant
        amountValue = targetSheet.Cells(currentRow, ���o�����z��).Value
        If Not IsNumeric(amountValue) Then
            Err.Raise 13, "���o�����S�̎擾", currentRow & "�s�ڂ�D��i���o�����z�j�����l�ł͂���܂���B"
        End If
        dataArray(i, 3) = CDbl(amountValue)
        
        ' E��F�c���i���l�`�F�b�N�j
        Dim balanceValue As Variant
        balanceValue = targetSheet.Cells(currentRow, �c����).Value
        If Not IsNumeric(balanceValue) Then
            Err.Raise 13, "���o�����S�̎擾", currentRow & "�s�ڂ�E��i�c���j�����l�ł͂���܂���B"
        End If
        dataArray(i, 4) = CDbl(balanceValue)
        
        ' F��F���ؒ��̖��ԍό������v�i���͂�����ΐ��l�`�F�b�N�j
        Dim principalValue As Variant
        principalValue = targetSheet.Cells(currentRow, ���ԍό�����).Value
        If principalValue <> "" Then
            If Not IsNumeric(principalValue) Then
                Err.Raise 13, "���o�����S�̎擾", currentRow & "�s�ڂ�F��i���ؒ��̖��ԍό������v�j�����l�ł͂���܂���B"
            End If
            dataArray(i, 5) = CDbl(principalValue)
        Else
            ' �󔒂̏ꍇ�͑O�̍s�̒l���g�p�A������i=1�̏ꍇ��0��ݒ�
            If i = 1 Then
                dataArray(i, 5) = 0
            Else
                dataArray(i, 5) = dataArray(i - 1, 5)
            End If
        End If
        
        ' dataArray(i, 5)�Z�o������̒ǉ������F�E�v���u�ԍϕ��v�ŏI���ꍇ�͓��o�����z�����炷
        If Len(dataArray(i, 2)) >= Len(�ԍϓE�v�ԍϕ�������) And Right(dataArray(i, 2), Len(�ԍϓE�v�ԍϕ�������)) = �ԍϓE�v�ԍϕ������� Then
            dataArray(i, 5) = dataArray(i, 5) - dataArray(i, 3)
        End If
        
NextRecord:
        currentRow = currentRow + 1
    Next i
    
    ' �����Ƃ��ă}�[�N���ꂽ���R�[�h���t�B���^�����O�i���t�������l�̃��R�[�h�j
    Dim validCount As Long
    validCount = 0
    For i = 1 To rowCount
        If dataArray(i, 1) <> ���t�����l Then
            validCount = validCount + 1
        End If
    Next i
    
    ' �L���ȃ��R�[�h���Ȃ��ꍇ�͋�z���Ԃ�
    If validCount = 0 Then
        ���o�����S�̎擾 = Array()
        Exit Function
    End If
    
    ' �t�B���^�����O��̔z����쐬
    Dim filteredArray() As Variant
    ReDim filteredArray(1 To validCount, 1 To 5)
    Dim validIndex As Long
    validIndex = 0
    
    For i = 1 To rowCount
        If dataArray(i, 1) <> ���t�����l Then
            validIndex = validIndex + 1
            For j = 1 To 5
                filteredArray(validIndex, j) = dataArray(i, j)
            Next j
        End If
    Next i
    
    ���o�����S�̎擾 = filteredArray
End Function

' ���������擾����֐�
' �w�肳�ꂽ�V�[�g��C��25�s�ڂ̃Z���l��Ԃ�
Public Function �������擾(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' �Z���l���擾
    cellValue = targetSheet.Cells(�������s, ��������).Value
    
    ' ���t�^���`�F�b�N
    If Not IsDate(cellValue) Then
        Err.Raise 13, "������", "�Z���l�����t�^�ł͂���܂���B"
    End If
    
    ' ���t�^�ɕϊ����ĕԂ�
    �������擾 = CDate(cellValue)
End Function

' �ؓ������擾�֐�
' �w�肳�ꂽ�V�[�g��29�s�ڂ���B�񂪋󔒂ɂȂ�܂ŁAB��i�ؓ������j��C��i�J�n���j�̃f�[�^���擾
Public Function �ؓ������擾(targetSheet As Worksheet) As Variant
    Dim ���ݍs As Long
    Dim ����() As Variant
    Dim �s�� As Long
    Dim i As Long
    
    ' �f�[�^�s�����J�E���g
    ���ݍs = �ؓ������J�n�s
    �s�� = 0
    
    Do While targetSheet.Cells(���ݍs, �ؓ�������).Value <> ""
        �s�� = �s�� + 1
        ���ݍs = ���ݍs + 1
    Loop
    
    ' �f�[�^�����݂��Ȃ��ꍇ�͋�̔z���Ԃ�
    If �s�� = 0 Then
        �ؓ������擾 = Array()
        Exit Function
    End If
    
    ' ���ʔz����������i�s�� x 2��j
    ReDim ����(1 To �s��, 1 To 2)
    
    ' �f�[�^���擾
    ���ݍs = �ؓ������J�n�s
    For i = 1 To �s��
        Dim �����l As Variant
        Dim �J�n���l As Variant
        
        ' B��i�ؓ������j���擾
        �����l = targetSheet.Cells(���ݍs, �ؓ�������).Value
        If Not IsNumeric(�����l) Then
            Err.Raise 13, "�ؓ������擾", "�ؓ����������l�^�ł͂���܂���B�s: " & ���ݍs
        End If
        ����(i, 1) = CDbl(�����l)
        
        ' C��i�J�n���j���擾
        �J�n���l = targetSheet.Cells(���ݍs, �ؓ������J�n����).Value
        If �J�n���l = "" Or IsEmpty(�J�n���l) Then
            ' �ŏ��̃��R�[�h�ŋ󔒂̏ꍇ�͍ŏ����t���Z�b�g
            If i = 1 Then
                ����(i, 2) = ���t�����l
            Else
                Err.Raise 13, "�ؓ������擾", "�J�n�����󔒂ł��B�s: " & ���ݍs
            End If
        Else
            If Not IsDate(�J�n���l) Then
                Err.Raise 13, "�ؓ������擾", "�J�n�������t�^�ł͂���܂���B�s: " & ���ݍs
            End If
            ����(i, 2) = CDate(�J�n���l)
        End If
        
        ���ݍs = ���ݍs + 1
    Next i
    
    �ؓ������擾 = ����
End Function

' �x�����Q�������擾�֐�
' �w�肳�ꂽ�V�[�g��15�s�ڂ���B�񂪋󔒂ɂȂ�܂ŁAB��i�x�����Q�������j��C��i�J�n���j�̃f�[�^���擾
Public Function �x�����Q�������擾(targetSheet As Worksheet) As Variant
    Dim ���ݍs As Long
    Dim ����() As Variant
    Dim �s�� As Long
    Dim i As Long
    
    ' �f�[�^�s�����J�E���g
    ���ݍs = �x�����Q�������J�n�s
    �s�� = 0
    
    Do While targetSheet.Cells(���ݍs, �x�����Q��������).Value <> ""
        �s�� = �s�� + 1
        ���ݍs = ���ݍs + 1
    Loop
    
    ' �f�[�^�����݂��Ȃ��ꍇ�͋�̔z���Ԃ�
    If �s�� = 0 Then
        �x�����Q�������擾 = Array()
        Exit Function
    End If
    
    ' ���ʔz����������i�s�� x 2��j
    ReDim ����(1 To �s��, 1 To 2)
    
    ' �f�[�^���擾
    ���ݍs = �x�����Q�������J�n�s
    For i = 1 To �s��
        Dim �����l As Variant
        Dim �J�n���l As Variant
        
        ' B��i�x�����Q�������j���擾
        �����l = targetSheet.Cells(���ݍs, �x�����Q��������).Value
        If Not IsNumeric(�����l) Then
            Err.Raise 13, "�x�����Q�������擾", "�x�����Q�����������l�^�ł͂���܂���B�s: " & ���ݍs
        End If
        ����(i, 1) = CDbl(�����l)
        
        ' C��i�J�n���j���擾
        �J�n���l = targetSheet.Cells(���ݍs, �x�����Q�������J�n����).Value
        If �J�n���l = "" Or IsEmpty(�J�n���l) Then
            ' �ŏ��̃��R�[�h�ŋ󔒂̏ꍇ�͍ŏ����t���Z�b�g
            If i = 1 Then
                ����(i, 2) = ���t�����l
            Else
                Err.Raise 13, "�x�����Q�������擾", "�J�n�����󔒂ł��B�s: " & ���ݍs
            End If
        Else
            If Not IsDate(�J�n���l) Then
                Err.Raise 13, "�x�����Q�������擾", "�J�n�������t�^�ł͂���܂���B�s: " & ���ݍs
            End If
            ����(i, 2) = CDate(�J�n���l)
        End If
        
        ���ݍs = ���ݍs + 1
    Next i
    
    �x�����Q�������擾 = ����
End Function

' �ԍϗ������擾�֐�
' 70�s�ڂ���B�񂪋󔒂ɂȂ�܂ŁAB��i���t�j�AC��i�E�v�j�AD��i�o�����z�j�̃f�[�^���擾
Public Function �ԍϗ������擾(targetSheet As Worksheet) As Variant
    Dim ���ݍs As Long
    Dim ����() As Variant
    Dim �s�� As Long
    Dim i As Long
    
    ' �f�[�^�s�����J�E���g
    ���ݍs = �ԍϗ����J�n�s
    �s�� = 0
    
    Do While targetSheet.Cells(���ݍs, �ԍϗ�����t��).Value <> ""
        �s�� = �s�� + 1
        ���ݍs = ���ݍs + 1
    Loop
    
    ' �f�[�^�����݂��Ȃ��ꍇ�͋�̔z���Ԃ�
    If �s�� = 0 Then
        �ԍϗ������擾 = Array()
        Exit Function
    End If
    
    ' ���ʔz����������i�s�� x 3��j
    ReDim ����(1 To �s��, 1 To 3)
    
    ' �f�[�^���擾
    ���ݍs = �ԍϗ����J�n�s
    For i = 1 To �s��
        ' B��i���t�j���擾
        Dim ���t�l As Variant
        ���t�l = targetSheet.Cells(���ݍs, �ԍϗ�����t��).Value
        If Not IsDate(���t�l) Then
            Err.Raise 13, "�ԍϗ������擾", ���ݍs & "�s�ڂ�B��i���t�j�����t�ł͂���܂���B"
        End If
        ����(i, 1) = CDate(���t�l)
        
        ' C��i�E�v�j���擾
        ����(i, 2) = CStr(targetSheet.Cells(���ݍs, �ԍϗ���E�v��).Value)
        
        ' D��i�o�����z�j���擾
        Dim �o�����z�l As Variant
        �o�����z�l = targetSheet.Cells(���ݍs, �ԍϗ����o�����z��).Value
        If Not IsNumeric(�o�����z�l) Then
            Err.Raise 13, "�ԍϗ������擾", ���ݍs & "�s�ڂ�D��i�o�����z�j�����l�ł͂���܂���B"
        End If
        ����(i, 3) = CDbl(�o�����z�l)
        
        ���ݍs = ���ݍs + 1
    Next i
    
    �ԍϗ������擾 = ����
End Function

' �v�Z���Ԃ̍ŏ������v�Z����֐�
Public Function �v�Z���ԍŏ����擾(targetSheet As Worksheet) As Date
    Dim �ԍϗ\��f�[�^ As Variant
    Dim �ŏ��� As Date
    Dim ���o���f�[�^ As Variant
    Dim i As Long
    Dim ���o���ŏ����t As Date
    Dim ���o���f�[�^���� As Boolean
    Dim �ԍϗ\��ŏ��� As Date
    
    ' �ԍϗ\������擾
    �ԍϗ\��f�[�^ = �ԍϗ\����擾(targetSheet)
    
    ' �ԍϗ\��f�[�^�����݂��A2�Ԗڂ̃��R�[�h������ꍇ�A2�Ԗڂ̓��t���擾
    If IsArray(�ԍϗ\��f�[�^) And UBound(�ԍϗ\��f�[�^, 1) >= 2 Then
        �ԍϗ\��ŏ��� = �ԍϗ\��f�[�^(2, 1) ' 2�Ԗڂ̕ԍϗ\���
    Else
        Err.Raise 13, "�v�Z���ԍŏ����擾", "�ԍϗ\�����2�Ԗڂ̃��R�[�h�����݂��܂���B"
    End If
    
    ' �ԍϗ\��ŏ����̑O����1���������l�Ƃ��čŏ����ɃZ�b�g
    �ŏ��� = DateSerial(Year(�ԍϗ\��ŏ���), Month(�ԍϗ\��ŏ���), 1)
    �ŏ��� = DateAdd("m", -1, �ŏ���)
    
    ' ���o�������擾
    ���o���f�[�^ = ���o�����S�̎擾(targetSheet)
    
    ' ���o���f�[�^�����݂��邩�`�F�b�N
    ���o���f�[�^���� = IsArray(���o���f�[�^) And UBound(���o���f�[�^, 1) > 0
    
    ' ���ԍϑO���f�[�^�Ƃ��Ďg�p
    Dim ���ԍϑO�� As Date
    ���ԍϑO�� = �ԍϗ\��f�[�^(1, 1)
    
    If ���o���f�[�^���� Then
        ' ���o�����̍ŏ����t���擾
        ���o���ŏ����t = ���o���f�[�^(1, 1) ' �ŏ��̓��t�ŏ�����
        For i = 2 To UBound(���o���f�[�^, 1)
            If ���o���f�[�^(i, 1) < ���o���ŏ����t Then
                ���o���ŏ����t = ���o���f�[�^(i, 1)
            End If
        Next i
        
        ' ���o�����ɂ��̍ŏ�����菬�������t�����邩�ǂ����m�F
        If ���o���ŏ����t < �ŏ��� Then
            ' ����΁A���̍ŏ�����Ԃ�
            �v�Z���ԍŏ����擾 = �ŏ���
            Exit Function
        End If
        
        ' �u���ԍϑO���v�����̍ŏ�����菬���������`�F�b�N
        If ���ԍϑO�� < �ŏ��� and ���ԍϑO�� > ���t�����l Then
            ' �ŏ�����Ԃ�
            �v�Z���ԍŏ����擾 = �ŏ���
            Exit Function
        End If
        
        ' ���ԍϑO�������t�����l�̏ꍇ�́A���o���ŏ����t��ݒ�
        If ���ԍϑO�� = ���t�����l Then
            �v�Z���ԍŏ����擾 = ���o���ŏ����t
        Else
            ' ��L�ȊO�̏ꍇ�́A���o�����̍ŏ����t�Ɓu���ԍϑO���v���r���āA�������ق���Ԃ�
            If ���o���ŏ����t < ���ԍϑO�� Then
                �v�Z���ԍŏ����擾 = ���o���ŏ����t
            Else
                �v�Z���ԍŏ����擾 = ���ԍϑO��
            End If
        End If
    Else
        ' ���o���f�[�^�����݂��Ȃ��ꍇ
        ' �u���ԍϑO���v�����̍ŏ�����菬���������`�F�b�N
        If ���ԍϑO�� < �ŏ��� and ���ԍϑO�� > ���t�����l Then
            ' �ŏ�����Ԃ�
            �v�Z���ԍŏ����擾 = �ŏ���
        Else
            ' ���ԍϑO����Ԃ�
            �v�Z���ԍŏ����擾 = ���ԍϑO��
        End If
    End If
End Function







