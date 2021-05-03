Attribute VB_Name = "security"
'#################################################
'参照設定：MicrosoftWbemScripting.SWbemLocator
'#################################################
Option Explicit

'#################################################
'個人情報を削除するマクロ
'#################################################

'==================================================
'workbookのBuiltinDocumentPropertiesを変更すると作成者情報とかが消える
'==================================================
Private Sub DelInfoTest()
    DelInfo ThisWorkbook
End Sub

Private Function DelInfo(ByVal wb As Workbook) As Boolean
    Dim vUserName As String
    
    On Error GoTo errhdl
    vUserName = Application.UserName
    
    'ユーザーネームは削除できないため、半角スペースを入れて実行させる必要がある
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
'プロパティを設定、削除するコード
'==================================================
Private Sub RemoveDocumentInformationSample()
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Sample12-1.xlsx")
    
    'エラー処理
    On Error Resume Next
    
    'テスト用の処理
    'タイトルを設定
    wb.BuiltinDocumentProperties("Title") = "Sample"
    
    '作成者をExcelUserに設定
    wb.BuiltinDocumentProperties("Author") = "ExcelUser"
    
    '最終保存者を"VBAUser"に設定
    wb.BuiltinDocumentProperties("LastAuthor") = "VBAUser"
    
    
   '処理を一旦停止
   Stop
   
   'プロパティ情報を削除
   wb.RemoveDocumentInformation xlRDIDocumentProperties
End Sub

'==================================================
'↑をFunctionにしたやつ
'引数_wb：文書情報をコントロールするワークブック
'==================================================
Private Function RemoveDocumentInformationSample2(ByVal wb As Workbook) As Boolean
    On Error GoTo errhdl
    
   'プロパティ情報を削除
   wb.RemoveDocumentInformation xlRDIDocumentProperties
   RemoveDocumentInformationSample2 = True
errhdl:
    MsgBox Err.Description, vbExclamation
End Function

'==================================================
'WMIを使ってセキュリティを上げる方法
'ここで参照設定WbemScripting.SWbemLocatorを使う
'これ使う？getusernameとかでいいんじゃないの？
'==================================================
Private Sub CheckLoginUser()
    Debug.Print IsCorrectLoginUser("HiroshiUrayama")
End Sub

Private Function IsCorrectLoginUser(ByVal UserName As String) As Boolean
    'WMIで使用するオブジェクト変数
    Dim oResult As Object
    Dim oTargetResult As Object
    Dim oLocator As Object
    Dim oService As Object
    Dim sMesStr As String
    
    'ローカルコンピューターに接続する
    Set oLocator = CreateObject("WbemScripting.SWbemLocator")
    Set oService = oLocator.ConnectServer
    
    'クエリー条件を指定する
    Set oResult = oService.ExecQuery("Select * From Win32_UserAccount")
    
    For Each oTargetResult In oResult
        If UserName = oTargetResult.Name Then
            IsCorrectLoginUser = True
            Exit For
        End If
    Next
    
End Function
