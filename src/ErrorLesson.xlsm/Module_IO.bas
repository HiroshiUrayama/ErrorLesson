Attribute VB_Name = "Module_IO"
'#################################################
'モジュールをすべてエクスポートするマクロ
'#################################################
Option Explicit

'==================================================
'モジュールをすべてエクスポートするマクロ
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
    
    'エクスポート先ディレクトリの取得
    ExportPath = ThisWorkbook.Path & "\Export_Modules"
    
    'エクスポート先がない場合、作成する
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
    MsgBox "処理が完了しました", vbInformation
End Sub

'==================================================
'モジュールをすべてインポートするマクロ
'==================================================
Private Sub ImpotrtModulesTest()
    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    ImportVBAModules wb
End Sub

Private Sub ImportVBAModules(ByVal wb As Workbook)
    Dim vPath As String
    Dim vFileName As String
    
    'インポートするモジュールのあるフォルダ
    vPath = ThisWorkbook.Path & "\Export_Modules\"
    vFileName = Dir(vPath & "*.*")
    
    Do While vFileName <> vbNullString
        If Right(vFileName, 4) = ".frx" Then GoTo continue
        wb.VBProject.VBComponents.Import vPath & vFileName
        vFileName = Dir
continue:
    Loop
    
    MsgBox "処理が完了しました", vbInformation
End Sub
