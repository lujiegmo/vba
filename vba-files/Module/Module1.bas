Attribute VB_Name = "Module1"
Option Explicit

' �萔��`

' ���o�����֘A�萔
Private Const ���o���J�n�s As Long = 46 ' ���o���f�[�^�J�n�s
Private Const ���o������ As Long = 2     ' B��F���o����
Private Const �E�v�� As Long = 3         ' C��F�E�v
Private Const ���o�����z�� As Long = 4   ' D��F���o�����z
Private Const �c���� As Long = 5         ' E��F�c��
Private Const ���ԍό����� As Long = 6 ' F��F���ؒ��̖��ԍό������v

' �������֘A�萔
Private Const �������� As Long = 3  ' C��
Private Const �������s As Long = 25  ' 25�s��

' �ؓ������֘A�萔
Private Const �ؓ������� As Long = 2  ' B��F�ؓ�����
Private Const �ؓ������J�n���� As Long = 3  ' C��F�J�n��
Private Const �ؓ������J�n�s As Long = 29  ' 29�s��

' �x�����Q�������֘A�萔
Private Const �x�����Q�������� As Long = 2  ' B��F�x�����Q������
Private Const �x�����Q�������J�n���� As Long = 3  ' C��F�J�n��
Private Const �x�����Q�������J�n�s As Long = 15  ' 15�s��

' �v�Z���쐬�p�X�֘A�萔
Private Const �v�Z���쐬�p�X�� As Long = 3  ' C��F�v�Z���쐬�p�X
Private Const �v�Z���쐬�p�X�s As Long = 7  ' 7�s��

' �o�͍��ڗ�萔�iA��`M��j
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

' �ԍϗ\����̒萔
Const �ԍϗ\��J�n�s As Long = 35
Const �ԍϗ\����� As Long = 3  ' C��
Const �ԍό����� As Long = 4    ' D��

' �v�Z���̍쐬�p�X�擾�֐�
' C��7�s�ڂ���v�Z���̍쐬�p�X���擾���A�󔒂̏ꍇ�̓G���[�𔭐�������
Public Function �v�Z���̍쐬�p�X�擾(targetSheet As Worksheet) As String
    Dim pathValue As Variant
    
    ' C��7�s�ڂ̒l���擾
    pathValue = targetSheet.Cells(�v�Z���쐬�p�X�s, �v�Z���쐬�p�X��).Value
    
    ' �󔒃`�F�b�N
    If pathValue = "" Or IsEmpty(pathValue) Then
        Err.Raise 13, "�v�Z���̍쐬�p�X�擾", "C��7�s�ځi�v�Z���̍쐬�p�X�j���󔒂ł��B"
    End If
    
    ' ������Ƃ��ĕԂ�
    �v�Z���̍쐬�p�X�擾 = CStr(pathValue)
End Function

' �ԍϗ\����擾�֐�
' 35�s�ڂ���J�n���AC�񂪋󔒂ɂȂ�܂ŕԍϗ\����A�ԍό����A�ԍό����݌v���擾
Public Function �ԍϗ\����擾(targetSheet As Worksheet) As Variant
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim �ԍό����݌v As Double
    
    currentRow = �ԍϗ\��J�n�s
    rowCount = 0
    
    ' �f�[�^�s�����J�E���g�iC�񂪋󔒂ɂȂ�܂Łj
    Do While targetSheet.Cells(currentRow, �ԍϗ\�����).Value <> ""
        rowCount = rowCount + 1
        currentRow = currentRow + 1
    Loop
    
    ' 35�s�ڂ�C�񂪋󔒂̏ꍇ��1�s�Ƃ��ď���
    If rowCount = 0 Then
        rowCount = 1
    End If
    
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
            dataArray(i, 1) = DateSerial(1900, 1, 1)
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

' ���핪�o�̓f�[�^�쐬�֐�
' �ԍϗ\�����2���R�[�h�ڂ��烋�[�v���ďo�̓f�[�^���쐬
Public Function ���핪�o�̓f�[�^�쐬(targetSheet As Worksheet) As Variant
    Dim �ԍϗ\��f�[�^ As Variant
    Dim ���o���f�[�^ As Variant
    Dim �ؓ������f�[�^ As Variant
    Dim �x�����Q�������f�[�^ As Variant
    Dim �v�Z���ԍŏ��� As Date
    Dim �o�͌���() As Variant
    Dim �o�͍s�� As Long
    Dim i As Long, j As Long
    
    ' 1. �ԍϗ\����擾
    �ԍϗ\��f�[�^ = �ԍϗ\����擾(targetSheet)
    ���o���f�[�^ = ���o�����擾(targetSheet)
    �ؓ������f�[�^ = �ؓ������擾(targetSheet)
    �x�����Q�������f�[�^ = �x�����Q�������擾(targetSheet)
    �v�Z���ԍŏ��� = �v�Z���ԍŏ����擾(targetSheet)
    
    ' �f�[�^���݃`�F�b�N
    If Not IsArray(�ԍϗ\��f�[�^) Or UBound(�ԍϗ\��f�[�^, 1) < 2 Then
        Err.Raise 13, "�o�̓f�[�^�쐬", "�ԍϗ\��f�[�^���s�����Ă��܂��B"
    End If
    
    ' �o�͌��ʔz��̏������i�ő�z��s���ŏ������j
    ReDim �o�͌���(1 To 1000, 1 To 13)
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
        
        ' ���������X�g�̏�����
        ReDim ���������X�g(1 To 100)
        �������� = 0
        
        ' 6. ���ԍϔN��1�̑O�������ԓ��ɂ��邩�`�F�b�N
        Dim ���ԍϔN��1�O�� As Date
        ���ԍϔN��1�O�� = DateSerial(Year(�ԍϗ\��O���f�[�^(0)), Month(�ԍϗ\��O���f�[�^(0)), 1)
        ���ԍϔN��1�O�� = DateAdd("m", -1, ���ԍϔN��1�O��)
        If ���ԍϔN��1�O�� >= ���ԊJ�n�� And ���ԍϔN��1�O�� <= ���ԏI���� Then
            �������� = �������� + 1
            ���������X�g(��������) = ���ԍϔN��1�O��
        End If
        
        ' 7. ���o�����̓��t�����ԓ��ɂ��邩�`�F�b�N
        If IsArray(���o���f�[�^) And UBound(���o���f�[�^, 1) > 0 Then
            For j = 1 To UBound(���o���f�[�^, 1)
                Dim ���o���� As Date
                ���o���� = ���o���f�[�^(j, 1)
                If ���o���� >= ���ԊJ�n�� And ���o���� <= ���ԏI���� Then
                    ' �����̕������Əd�����Ȃ����`�F�b�N
                    Dim �d���t���O As Boolean
                    �d���t���O = False
                    Dim k As Long
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
            Next j
        End If
        
        ' 8. �ؓ������f�[�^�̊J�n�������ԓ��ɂ��邩�`�F�b�N
        If IsArray(�ؓ������f�[�^) And UBound(�ؓ������f�[�^, 1) > 0 Then
            For j = 1 To UBound(�ؓ������f�[�^, 1)
                Dim �ؓ������J�n�� As Date
                �ؓ������J�n�� = �ؓ������f�[�^(j, 2)
                If �ؓ������J�n�� >= ���ԊJ�n�� And �ؓ������J�n�� <= ���ԏI���� Then
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
                If �x�����Q�������J�n�� >= ���ԊJ�n�� And �x�����Q�������J�n�� <= ���ԏI���� Then
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
            �o�͌���(�o�͍s��, �o��_�X�e�[�^�X��) = "����"
            
            ' �C�x���g
            �o�͌���(�o�͍s��, �o��_�C�x���g��) = "���ԍ�"
            
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
            �o�͌���(�o�͍s��, �o��_�v�Z������) = DateDiff("d", �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����), �o�͌���(�o�͍s��, �o��_�v�Z���ԏI������)) + 1
            
            ' �Ώی����̌v�Z
            Dim �Ώی��� As Double
            Dim �c�� As Double
            Dim ���ؒ����ԍό��� As Double
            
            ' ���o���f�[�^����c���Ɖ��ؒ����ԍό������擾
            �c�� = 0
            ���ؒ����ԍό��� = 0
            If IsArray(���o���f�[�^) And UBound(���o���f�[�^, 1) > 0 Then
                �c�� = ���o���f�[�^(UBound(���o���f�[�^, 1), 4) ' �Ō�̎c��
                ���ؒ����ԍό��� = ���o���f�[�^(UBound(���o���f�[�^, 1), 5) ' �Ō�̉��ؒ����ԍό���
            End If
            
            �Ώی��� = �c�� - ���ؒ����ԍό���
            
            ' �v�Z���ԊJ�n�����ԍϗ\��O���f�[�^�̓��t�Ɠ��������傫���ꍇ
            If �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) >= �ԍϗ\��O���f�[�^(0) Then
                �Ώی��� = �Ώی��� - �ԍϗ\��O���f�[�^(2) ' �ԍό����݌v�����炷
            End If
            
            �o�͌���(�o�͍s��, �o��_�Ώی�����) = �Ώی���
            
            ' �����̎擾
            Dim ���� As Double
            ���� = 0
            If IsArray(�ؓ������f�[�^) And UBound(�ؓ������f�[�^, 1) > 0 Then
                For k = 1 To UBound(�ؓ������f�[�^, 1)
                    If �ؓ������f�[�^(k, 2) = �o�͌���(�o�͍s��, �o��_�v�Z���ԊJ�n����) Then
                        ���� = �ؓ������f�[�^(k, 1)
                        Exit For
                    End If
                Next k
            End If
            �o�͌���(�o�͍s��, �o��_������) = ����
            
            ' �ϐ��A�������z�A�x�����Q���͐ݒ�s�iExcel��������j
            �o�͌���(�o�͍s��, �o��_�ϐ���) = ""
            �o�͌���(�o�͍s��, �o��_�������z��) = ""
            �o�͌���(�o�͍s��, �o��_�x�����Q����) = ""
        Next j
    Next i
    
    ' ���ʔz��̃T�C�Y�𒲐�
    If �o�͍s�� > 0 Then
        ReDim Preserve �o�͌���(1 To �o�͍s��, 1 To 13)
        �o�̓f�[�^�쐬 = �o�͌���
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

' ���o�����擾�֐�
Public Function ���o�����擾(targetSheet As Worksheet) As Variant
    Dim startRow As Long
    Dim currentRow As Long
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim i As Long
    
    startRow = ���o���J�n�s ' �J�n�s
    currentRow = startRow
    rowCount = 0
    
    ' �f�[�^�s�����J�E���g�iB�񂪋󔒂ɂȂ�܂ŁA�E�v���u�ԍϕ��v�ŏI���s�͏��O�j
    Do While targetSheet.Cells(currentRow, ���o������).Value <> ""
        Dim �E�v�l As String
        �E�v�l = CStr(targetSheet.Cells(currentRow, �E�v��).Value)
        ' �E�v���u�ԍϕ��v�ŏI���Ȃ��ꍇ�̂݃J�E���g
        If Not (Len(�E�v�l) >= 3 And Right(�E�v�l, 3) = "�ԍϕ�") Then
            rowCount = rowCount + 1
        End If
        currentRow = currentRow + 1
    Loop
    
    ' �f�[�^�����݂��Ȃ��ꍇ�͋�̔z���Ԃ�
    If rowCount = 0 Then
        ���o�����擾 = Array()
        Exit Function
    End If
    
    ' �z����������i�s�� x 5��j
    ReDim dataArray(1 To rowCount, 1 To 5)
    
    ' �f�[�^���擾���ăo���f�[�V����
    Dim arrayIndex As Long
    arrayIndex = 1
    currentRow = startRow
    
    Do While targetSheet.Cells(currentRow, ���o������).Value <> ""
        Dim �E�v�l As String
        �E�v�l = CStr(targetSheet.Cells(currentRow, �E�v��).Value)
        
        ' �E�v���u�ԍϕ��v�ŏI���Ȃ��ꍇ�̂ݏ���
        If Not (Len(�E�v�l) >= 3 And Right(�E�v�l, 3) = "�ԍϕ�") Then
            ' B��F���o�����i���t�`�F�b�N�j
            Dim dateValue As Variant
            dateValue = targetSheet.Cells(currentRow, ���o������).Value
            If Not IsDate(dateValue) Then
                Err.Raise 13, "���o�����擾", currentRow & "�s�ڂ�B��i���o�����j�����t�ł͂���܂���B"
            End If
            dataArray(arrayIndex, 1) = CDate(dateValue)
            
            ' C��F�E�v�i������A�`�F�b�N�s�v�j
            dataArray(arrayIndex, 2) = �E�v�l
            
            ' D��F���o�����z�i���l�`�F�b�N�j
            Dim amountValue As Variant
            amountValue = targetSheet.Cells(currentRow, ���o�����z��).Value
            If Not IsNumeric(amountValue) Then
                Err.Raise 13, "���o�����擾", currentRow & "�s�ڂ�D��i���o�����z�j�����l�ł͂���܂���B"
            End If
            dataArray(arrayIndex, 3) = CDbl(amountValue)
            
            ' E��F�c���i���l�`�F�b�N�j
            Dim balanceValue As Variant
            balanceValue = targetSheet.Cells(currentRow, �c����).Value
            If Not IsNumeric(balanceValue) Then
                Err.Raise 13, "���o�����擾", currentRow & "�s�ڂ�E��i�c���j�����l�ł͂���܂���B"
            End If
            dataArray(arrayIndex, 4) = CDbl(balanceValue)
            
            ' F��F���ؒ��̖��ԍό������v�i���͂�����ΐ��l�`�F�b�N�j
            Dim principalValue As Variant
            principalValue = targetSheet.Cells(currentRow, ���ԍό�����).Value
            If principalValue <> "" Then
                If Not IsNumeric(principalValue) Then
                    Err.Raise 13, "���o�����擾", currentRow & "�s�ڂ�F��i���ؒ��̖��ԍό������v�j�����l�ł͂���܂���B"
                End If
                dataArray(arrayIndex, 5) = CDbl(principalValue)
            Else
                ' �󔒂̏ꍇ�͑O�̍s�̒l���g�p�A������arrayIndex=1�̏ꍇ��0��ݒ�
                If arrayIndex = 1 Then
                    dataArray(arrayIndex, 5) = 0
                Else
                    dataArray(arrayIndex, 5) = dataArray(arrayIndex - 1, 5)
                End If
            End If
            
            arrayIndex = arrayIndex + 1
        End If
        
        currentRow = currentRow + 1
    Loop
    
    ���o�����擾 = dataArray
End Function

' ���������擾����֐�
' �w�肳�ꂽ�V�[�g��C��25�s�ڂ̃Z���l��Ԃ�
Public Function ������(targetSheet As Worksheet) As Date
    Dim cellValue As Variant
    
    ' �Z���l���擾
    cellValue = targetSheet.Cells(�������s, ��������).Value
    
    ' ���t�^���`�F�b�N
    If Not IsDate(cellValue) Then
        Err.Raise 13, "������", "�Z���l�����t�^�ł͂���܂���B"
    End If
    
    ' ���t�^�ɕϊ����ĕԂ�
    ������ = CDate(cellValue)
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
                ����(i, 2) = DateSerial(1900, 1, 1)
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
                ����(i, 2) = DateSerial(1900, 1, 1)
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
    ���o���f�[�^ = ���o�����擾(targetSheet)
    
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
        If ���ԍϑO�� < �ŏ��� Then
            ' �ŏ�����Ԃ�
            �v�Z���ԍŏ����擾 = �ŏ���
            Exit Function
        End If
        
        ' ��L�ȊO�̏ꍇ�́A���o�����̍ŏ����t�Ɓu���ԍϑO���v���r���āA�������ق���Ԃ�
        If ���o���ŏ����t < ���ԍϑO�� Then
            �v�Z���ԍŏ����擾 = ���o���ŏ����t
        Else
            �v�Z���ԍŏ����擾 = ���ԍϑO��
        End If
    Else
        ' ���o���f�[�^�����݂��Ȃ��ꍇ
        ' �u���ԍϑO���v�����̍ŏ�����菬���������`�F�b�N
        If ���ԍϑO�� < �ŏ��� Then
            ' �ŏ�����Ԃ�
            �v�Z���ԍŏ����擾 = �ŏ���
        Else
            ' ���ԍϑO����Ԃ�
            �v�Z���ԍŏ����擾 = ���ԍϑO��
        End If
    End If
End Function
