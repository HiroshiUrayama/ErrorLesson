Attribute VB_Name = "Err_One"
Option Explicit

'#################################################
'���僌�x���ł͌����đ���Ȃ�ExcelVBA�����̂��߂̋Z�p
'#################################################

'==================================================
'�G���[�Ώ��̃��b�X��
'==================================================

'---------------------------------------------------
'�u�b�N��V�K�쐬���Ė��O�����ĕۑ�����
'---------------------------------------------------
Private Sub SaveWorkbooksSample()
    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    
    '���ꖼ�̂̃��[�N�u�b�N�����݂���ꍇ�A�����Ń_�C�A���O�m�F���o��
    '�͂�����������ꍇ�͂������A����������������ƃG���[�ɂȂ�(���s���G���[1004)
    wb.SaveAs ThisWorkbook.Path & "\Sample12-1.xlsx"
    
    wb.Close
End Sub

'---------------------------------------------------
'�G���[���Ɗ��m�����ꍇ�A���ɏ����ł��낤�R�[�h�c���щz���Ď����ōl�����R�[�h
'---------------------------------------------------
Private Sub SaveWorkbooksSample2()
    Dim wb As Workbook
    Dim fso As Object:    Set fso = CreateObject("scripting.FileSystemObject")
    Dim fileName As String:    fileName = ThisWorkbook.Path & "\Sample12-1.xlsx"

    '�u�b�N�쐬����
    Set wb = Workbooks.Add
    
    '---------------------------------------------------
    '�t�@�C����鏈��������O�ɁA���S�Ƀ`�F�b�N�����{���č쐬���̂��~�߂�
    '---------------------------------------------------
    '�t�@�C�������݂��Ă�����A������u�b�N����ď������I����
    If fso.fileexists(fileName) Then wb.Close: Exit Sub
    '�u�b�N���Ȃ��ꍇ�̓Z�[�u���ĕ���
    wb.SaveAs fileName
    wb.Close
End Sub

'---------------------------------------------------
'����
'---------------------------------------------------
Private Sub SaveWrokbookSample3()
    Dim wb As Workbook
    Dim vFileName As String: vFileName = "Sample12-1.xlsx"
    '---------------------------------------------------
    '�����̓`�F�b�N
    '---------------------------------------------------
    'Dir�֐��́A�t�H���_�̒��Ƀt�@�C�������݂��Ă���ꍇ�����t�@�C�������E���A����������
    If Len(Dir(ThisWorkbook.Path & "\" & vFileName)) <> 0 Then
        If MsgBox("�����̃u�b�N�����łɑ��݂��܂��B�㏑���ۑ����܂����H", vbYesNo + vbExclamation) = vbNo Then
            '�����ŏ������I�����Ă��邪�A�ʖ��ۑ��ɂ���Ƃ��A���@�͐F�X����
            Exit Sub
            
        End If
    End If
    '---------------------------------------------------
    '�������珈��
    '---------------------------------------------------
    '�u�b�N�V�K�ǉ�
    Set wb = Workbooks.Add
    Application.DisplayAlerts = False
    wb.SaveAs ThisWorkbook.Path & "\" & vFileName
    Application.DisplayAlerts = True
    wb.Close
End Sub

'---------------------------------------------------
'���̃u�b�N�̓��e���擾����(��x�u�b�N���J���Ă��̃u�b�N���擾������)
'---------------------------------------------------
Private Sub OpenLinkedBook()
    Dim wb As Workbook
    
    '���ʂɊJ���R�[�h�����ǁA�t�@�C���̕����������o�����������u�b�N�̎Q�Ǝ����܂�ł���ƍX�V���܂����H�\�����o�Ď~�܂�
    Set wb = Workbooks.Open(fileName:=ThisWorkbook.Path & "\Sample12-1.xlsx")

    MsgBox wb.Worksheets(1).Range("A1").Value
End Sub

'---------------------------------------------------
'�����N���X�V�����Ƀu�b�N���J���R�[�h�������Ă���
'---------------------------------------------------
Private Sub OpenLinkedBook2()
    Dim wb As Workbook
    
    '���ʂɊJ���R�[�h�����ǁA�t�@�C���̕����������o�����������u�b�N�̎Q�Ǝ����܂�ł���ƍX�V���܂����H�\�����o�Ď~�܂�
    'updatelinks�p�����[�^��ݒ肵�A���������N���ݒ肳��Ă����Ƃ��Ă����̕\�����o�Ȃ��悤�ɐݒ肵�Ă������Ƃ��d�v
    Set wb = Workbooks.Open(fileName:=ThisWorkbook.Path & "\Sample12-1.xlsx", UpdateLinks:=0)

    MsgBox wb.Worksheets(1).Range("A1").Value
End Sub

'==================================================
'����Funciton���Ăяo��
'==================================================
Private Sub ResetSheetTest()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(1)
    
    ResetSheet sh
End Sub

'---------------------------------------------------
'���ׂĂ̍s�Ɨ��\������R�[�h
'���������A�I�[�g�t�B���^�[���\���񂪑��݂��Ă���\��������̂ŁA�ŏ��ɂ������S���������Ă���
'---------------------------------------------------
Private Function ResetSheet(ByVal sh As Worksheet) As Boolean
    On Error GoTo errhdl
    With sh
        If .FilterMode Then
            .ShowAllData
        End If
        
        .Outline.ShowLevels columnlevels:=5
        .Outline.ShowLevels rowlevels:=5
        
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With
    
    ResetSheet = True
    Exit Function
    
errhdl:
    MsgBox Err.Description, vbExclamation
End Function

'�d�l���ŏ����猈�߂��Ƃ��Ă��A���ۂ̉^�p���ɂ͑z��O�̗l�X�ȃP�[�X����������B
'���O�ɑz��O��z�肵�ăv���O��������邱�Ƃ���؁B

'�����܂ł�肫���Ă��܂��[���ɂȂ�Ȃ��̂��G���[�B
'On Error goto �X�e�[�g�����g�̃n�C���x���Ȏg�����A�g���ׂ��P�[�X���o����B
'==================================================

