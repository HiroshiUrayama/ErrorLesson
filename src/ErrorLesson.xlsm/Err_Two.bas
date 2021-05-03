Attribute VB_Name = "Err_Two"
Option Explicit

'#################################################
'���ݍ��񂾃G���[����
'���z��ς��邱�Ƃ���؁B�u�G���[�𔭐������Ȃ��悤�Ɂv�ł͂Ȃ��A�u�G���[�������I�ɗ��p���ăR�[�h���X�b�L��������v
'#################################################

'---------------------------------------------------
'�����Ăяo���v���V�[�W��
'---------------------------------------------------
Private Sub SetNewSheetName()
    If IsCorrectSheetName("AAA") Then
        MsgBox "���̃V�[�g���͗L���ł�", vbInformation
    Else
        MsgBox "���̃V�[�g���͖����ł�", vbExclamation
    End If
End Sub

'---------------------------------------------------
'�V�[�g�����L�����ǂ����𔻒肷��v���V�[�W��
'---------------------------------------------------
Public Function IsCorrectSheetName(ByVal vName As String) As Boolean
    Dim sh As Worksheet
    
    '�G���[�������J�n
    On Error Resume Next
    
    '�V�[�g���쐬����
    Set sh = Worksheets.Add
    
    '�V�[�g��������
    sh.Name = vName
    
    '�G���[�ԍ���0��(����I��)�ȊO��������^�G���[���g���ق����������肵�ď�����
    If Err.Number = 0 Then
        '�������Ȃ�=�L��
        IsCorrectSheetName = True
    Else
        '��������=����
        IsCorrectSheetName = False
    End If
    On Error GoTo 0
    
    '---------------------------------------------------
    '������G���[���N�������ɋ�ʂ��悤�Ǝv���ƁA��ρc
        '���u�����N�ɂȂ��Ă��Ȃ���
        '�����[�N�V�[�g�Ɏg���Ȃ�����(*��!�Ȃ�)�������Ă��Ȃ���
        '�������̃��[�N�V�[�g���Ɠ���(�啶���E�������A���p�E�S�p�̋�ʂȂ��j�ł͂Ȃ���
        '����������31�����ȓ��ɂȂ��Ă��邩
        '���������S�����肵�Ȃ��Ƃ����Ȃ��I�I�I
        '�R�[�h�������Ȃ�c
    '---------------------------------------------------
    
    '�ǉ��������[�N�V�[�g���폜����
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True

End Function

'---------------------------------------------------
'On Error Resume Next(�͈͂�����I�Ɏg��)
'---------------------------------------------------
Private Sub OnErrorResumeNextSample()
    On Error Resume Next
    
    '�G���[����������\���̂��鏈��
    '�G���[�̗L���̃`�F�b�N
    If Err.Number <> 0 Then
        '�G���[����
    End If
    '�G���[�����̏I��
    On Error GoTo 0
    
    '���̑��̏���
    
End Sub

'---------------------------------------------------
'On Error goto�X�e�[�g�����g(��{)
'---------------------------------------------------
Private Sub OnErrorResumeNextSample1()
        
    '�G���[�����̊J�n
    On Error GoTo errhdl
    '�G���[����������\���̂��鏈��
    '���̑��̏���
    '�v���V�[�W���̏I��
    Exit Sub

errhdl:
    '�G���[����
End Sub

'---------------------------------------------------
'On Error goto�X�e�[�g�����g(�r���ŃG���[����������ꍇ�̑Ώ�)
'---------------------------------------------------
Private Sub OnErrorGotoSample2()
    '�G���[�����̊J�n
    On Error GoTo errhdl
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "Sample12-1.xlsx")
    
    '���̑��̏���
    Exit Sub
    
errhdl:
    '���̏ꍇ���ƁA�J�����u�b�N�͂��̂܂ܕ��u����Ă��܂�
    MsgBox Err.Description, vbExclamation
End Sub

'---------------------------------------------------
'On Error goto�X�e�[�g�����g(���P������)
'---------------------------------------------------
Private Sub OnErrorGotoSample3()
    '�G���[�����̊J�n
    On Error GoTo errhdl
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Sample12-1.xlsx")
    
    '���̑��̏���
    
    'Exit�����ɏI������������
Exithdl:
    On Error Resume Next
    wb.Close
    On Error GoTo 0
    Exit Sub
    
errhdl:
    '���̏ꍇ���ƁA�J�����u�b�N�͂��̂܂ܕ��u����Ă��܂�
    MsgBox Err.Description, vbExclamation
    Resume Exithdl
End Sub

'#################################################
'�G���[���𗘗p���邾���łȂ��āA�X�ɓƎ��̃G���[�𔭐������ăv���O�����𐧌䂷��
'#################################################

'---------------------------------------------------
'����G���[�𔭐�������R�[�h
'---------------------------------------------------
Private Sub CheckScoreTest()
    MsgBox CheckScore(-100)
End Sub

Private Function CheckScore(ByVal num As Long) As Variant
    On Error GoTo errhdl
    Select Case num
        Case 0 To 49
            CheckScore = "�����NC"
        Case 50 To 79
           CheckScore = "�����NC"
        Case 80 To 100
            CheckScore = "�����NC"
        Case Else
            Err.Raise 1000
        End Select
        Exit Function
errhdl:
    CheckScore = "�l������������܂���"
End Function
