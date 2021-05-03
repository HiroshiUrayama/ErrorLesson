Attribute VB_Name = "security"
'#################################################
'�Q�Ɛݒ�FMicrosoftWbemScripting.SWbemLocator
'#################################################
Option Explicit

'#################################################
'�l�����폜����}�N��
'#################################################

'==================================================
'workbook��BuiltinDocumentProperties��ύX����ƍ쐬�ҏ��Ƃ���������
'==================================================
Private Sub DelInfoTest()
    DelInfo ThisWorkbook
End Sub

Private Function DelInfo(ByVal wb As Workbook) As Boolean
    Dim vUserName As String
    
    On Error GoTo errhdl
    vUserName = Application.UserName
    
    '���[�U�[�l�[���͍폜�ł��Ȃ����߁A���p�X�y�[�X�����Ď��s������K�v������
    Application.UserName = " "
    
    With wb
        With .BuiltinDocumentProperties
            .Item("Author").Value = Empty
            .Item("Company").Value = Empty
            .Item("Manager").Value = Empty
        End With
        .Save
    End With

    Application.UserName = vUserName
    DelInfo = True
    Exit Function
    
    
errhdl:
    MsgBox Err.Description, vbExclamation
End Function

'==================================================
'�v���p�e�B��ݒ�A�폜����R�[�h
'==================================================
Private Sub RemoveDocumentInformationSample()
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Sample12-1.xlsx")
    
    '�G���[����
    On Error Resume Next
    
    '�e�X�g�p�̏���
    '�^�C�g����ݒ�
    wb.BuiltinDocumentProperties("Title") = "Sample"
    
    '�쐬�҂�ExcelUser�ɐݒ�
    wb.BuiltinDocumentProperties("Author") = "ExcelUser"
    
    '�ŏI�ۑ��҂�"VBAUser"�ɐݒ�
    wb.BuiltinDocumentProperties("LastAuthor") = "VBAUser"
    
    
   '��������U��~
   Stop
   
   '�v���p�e�B�����폜
   wb.RemoveDocumentInformation xlRDIDocumentProperties
End Sub

'==================================================
'����Function�ɂ������
'����_wb�F���������R���g���[�����郏�[�N�u�b�N
'==================================================
Private Function RemoveDocumentInformationSample2(ByVal wb As Workbook) As Boolean
    On Error GoTo errhdl
    
   '�v���p�e�B�����폜
   wb.RemoveDocumentInformation xlRDIDocumentProperties
   RemoveDocumentInformationSample2 = True
errhdl:
    MsgBox Err.Description, vbExclamation
End Function

'==================================================
'WMI���g���ăZ�L�����e�B���グ����@
'�����ŎQ�Ɛݒ�WbemScripting.SWbemLocator���g��
'����g���Hgetusername�Ƃ��ł����񂶂�Ȃ��́H
'==================================================
Private Sub CheckLoginUser()
    Debug.Print IsCorrectLoginUser("HiroshiUrayama")
End Sub

Private Function IsCorrectLoginUser(ByVal UserName As String) As Boolean
    'WMI�Ŏg�p����I�u�W�F�N�g�ϐ�
    Dim oResult As Object
    Dim oTargetResult As Object
    Dim oLocator As Object
    Dim oService As Object
    Dim sMesStr As String
    
    '���[�J���R���s���[�^�[�ɐڑ�����
    Set oLocator = CreateObject("WbemScripting.SWbemLocator")
    Set oService = oLocator.ConnectServer
    
    '�N�G���[�������w�肷��
    Set oResult = oService.ExecQuery("Select * From Win32_UserAccount")
    
    For Each oTargetResult In oResult
        If UserName = oTargetResult.Name Then
            IsCorrectLoginUser = True
            Exit For
        End If
    Next
    
End Function
