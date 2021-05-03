Attribute VB_Name = "Err_Two"
Option Explicit

'#################################################
'踏み込んだエラー処理
'発想を変えることが大切。「エラーを発生させないように」ではなく、「エラーを効率的に利用してコードをスッキリさせる」
'#################################################

'---------------------------------------------------
'↓を呼び出すプロシージャ
'---------------------------------------------------
Private Sub SetNewSheetName()
    If IsCorrectSheetName("AAA") Then
        MsgBox "このシート名は有効です", vbInformation
    Else
        MsgBox "このシート名は無効です", vbExclamation
    End If
End Sub

'---------------------------------------------------
'シート名が有効かどうかを判定するプロシージャ
'---------------------------------------------------
Public Function IsCorrectSheetName(ByVal vName As String) As Boolean
    Dim sh As Worksheet
    
    'エラー処理を開始
    On Error Resume Next
    
    'シートを作成する
    Set sh = Worksheets.Add
    
    'シート名をつける
    sh.Name = vName
    
    'エラー番号が0番(正常終了)以外だったら／エラーを使うほうがあっさりして書ける
    If Err.Number = 0 Then
        '発生しない=有効
        IsCorrectSheetName = True
    Else
        '発生する=無効
        IsCorrectSheetName = False
    End If
    On Error GoTo 0
    
    '---------------------------------------------------
    'これをエラーを起こさずに区別しようと思うと、大変…
        '※ブランクになっていないか
        '※ワークシートに使えない文字(*や!など)が入っていないか
        '※既存のワークシート名と同名(大文字・小文字、半角・全角の区別なし）ではないか
        '※文字数が31文字以内になっているか
        '↑こいつらを全部判定しないといけない！！！
        'コードが長くなる…
    '---------------------------------------------------
    
    '追加したワークシートを削除する
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True

End Function

'---------------------------------------------------
'On Error Resume Next(範囲を限定的に使う)
'---------------------------------------------------
Private Sub OnErrorResumeNextSample()
    On Error Resume Next
    
    'エラーが発生する可能性のある処理
    'エラーの有無のチェック
    If Err.Number <> 0 Then
        'エラー処理
    End If
    'エラー処理の終了
    On Error GoTo 0
    
    'その他の処理
    
End Sub

'---------------------------------------------------
'On Error gotoステートメント(基本)
'---------------------------------------------------
Private Sub OnErrorResumeNextSample1()
        
    'エラー処理の開始
    On Error GoTo errhdl
    'エラーが発生する可能性のある処理
    'その他の処理
    'プロシージャの終了
    Exit Sub

errhdl:
    'エラー処理
End Sub

'---------------------------------------------------
'On Error gotoステートメント(途中でエラーが発生する場合の対処)
'---------------------------------------------------
Private Sub OnErrorGotoSample2()
    'エラー処理の開始
    On Error GoTo errhdl
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "Sample12-1.xlsx")
    
    'その他の処理
    Exit Sub
    
errhdl:
    'この場合だと、開いたブックはそのまま放置されてしまう
    MsgBox Err.Description, vbExclamation
End Sub

'---------------------------------------------------
'On Error gotoステートメント(改善した例)
'---------------------------------------------------
Private Sub OnErrorGotoSample3()
    'エラー処理の開始
    On Error GoTo errhdl
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Sample12-1.xlsx")
    
    'その他の処理
    
    'Exitせずに終了処理をする
Exithdl:
    On Error Resume Next
    wb.Close
    On Error GoTo 0
    Exit Sub
    
errhdl:
    'この場合だと、開いたブックはそのまま放置されてしまう
    MsgBox Err.Description, vbExclamation
    Resume Exithdl
End Sub

'#################################################
'エラー情報を利用するだけでなくて、更に独自のエラーを発生させてプログラムを制御する
'#################################################

'---------------------------------------------------
'自らエラーを発生させるコード
'---------------------------------------------------
Private Sub CheckScoreTest()
    MsgBox CheckScore(-100)
End Sub

Private Function CheckScore(ByVal num As Long) As Variant
    On Error GoTo errhdl
    Select Case num
        Case 0 To 49
            CheckScore = "ランクC"
        Case 50 To 79
           CheckScore = "ランクC"
        Case 80 To 100
            CheckScore = "ランクC"
        Case Else
            Err.Raise 1000
        End Select
        Exit Function
errhdl:
    CheckScore = "値が正しくありません"
End Function
