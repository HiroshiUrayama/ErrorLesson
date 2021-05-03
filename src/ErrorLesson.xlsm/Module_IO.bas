Attribute VB_Name = "Module_IO"
'#################################################
'���W���[�������ׂăG�N�X�|�[�g����}�N��
'#################################################
Option Explicit

'==================================================
'���W���[�������ׂăG�N�X�|�[�g����}�N��
'==================================================
Private Enum ComponentType
    eStandard = 1
    eclass = 2
    euserform = 3
    eexcelobjects = 100
End Enum

Private Sub ExportVBAModules()
    Dim TempComponent As Object
    Dim ExportPath As String
    
    '�G�N�X�|�[�g��f�B���N�g���̎擾
    ExportPath = ThisWorkbook.Path & "\Export_Modules"
    
    '�G�N�X�|�[�g�悪�Ȃ��ꍇ�A�쐬����
    If Len(Dir(ExportPath, vbDirectory)) = 0 Then
        MkDir ExportPath
    End If
    
    For Each TempComponent In ThisWorkbook.VBProject.VBComponents
        Select Case TempComponent.Type
            Case ComponentType.eStandard
                TempComponent.Export ExportPath & "\" & TempComponent.Name & ".bas"
            Case ComponentType.eclass
                TempComponent.Export ExportPath & "\" & TempComponent.Name & ".cls"
            Case ComponentType.euserform
                TempComponent.Export ExportPath & "\" & TempComponent.Name & ".frm"
            Case ComponentType.eexcelobjects
                TempComponent.Export ExportPath & "\" & TempComponent.Name & ".cls"
            Case Else
        End Select
    Next
    MsgBox "�������������܂���", vbInformation
End Sub

'==================================================
'���W���[�������ׂăC���|�[�g����}�N��
'==================================================
Private Sub ImpotrtModulesTest()
    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    ImportVBAModules wb
End Sub

Private Sub ImportVBAModules(ByVal wb As Workbook)
    Dim vPath As String
    Dim vFileName As String
    
    '�C���|�[�g���郂�W���[���̂���t�H���_
    vPath = ThisWorkbook.Path & "\Export_Modules\"
    vFileName = Dir(vPath & "*.*")
    
    Do While vFileName <> vbNullString
        If Right(vFileName, 4) = ".frx" Then GoTo continue
        wb.VBProject.VBComponents.Import vPath & vFileName
        vFileName = Dir
continue:
    Loop
    
    MsgBox "�������������܂���", vbInformation
End Sub
