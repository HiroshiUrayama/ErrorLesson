Attribute VB_Name = "Err_One"
Option Explicit

'#################################################
'入門レベルでは決して足りないExcelVBA実務のための技術
'#################################################

'==================================================
'エラー対処のレッスン
'==================================================

'---------------------------------------------------
'ブックを新規作成して名前をつけて保存する
'---------------------------------------------------
Private Sub SaveWorkbooksSample()
    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    
    '同一名称のワークブックが存在する場合、ここでダイアログ確認が出る
    'はいを押下する場合はいいが、いいえを押下するとエラーになる(実行時エラー1004)
    wb.SaveAs ThisWorkbook.Path & "\Sample12-1.xlsx"
    
    wb.Close
End Sub

'---------------------------------------------------
'エラーだと感知した場合、次に書くであろうコード…を飛び越して自分で考えたコード
'---------------------------------------------------
Private Sub SaveWorkbooksSample2()
    Dim wb As Workbook
    Dim fso As Object:    Set fso = CreateObject("scripting.FileSystemObject")
    Dim fileName As String:    fileName = ThisWorkbook.Path & "\Sample12-1.xlsx"

    'ブック作成する
    Set wb = Workbooks.Add
    
    '---------------------------------------------------
    'ファイル作る処理をする前に、完全にチェックを実施して作成自体を止める
    '---------------------------------------------------
    'ファイルが存在していたら、作ったブックを閉じて処理を終える
    If fso.fileexists(fileName) Then wb.Close: Exit Sub
    'ブックがない場合はセーブして閉じる
    wb.SaveAs fileName
    wb.Close
End Sub

'---------------------------------------------------
'答え
'---------------------------------------------------
Private Sub SaveWrokbookSample3()
    Dim wb As Workbook
    Dim vFileName As String: vFileName = "Sample12-1.xlsx"
    '---------------------------------------------------
    'ここはチェック
    '---------------------------------------------------
    'Dir関数は、フォルダの中にファイルが存在している場合だけファイル名を拾う、だったかな
    If Len(Dir(ThisWorkbook.Path & "\" & vFileName)) <> 0 Then
        If MsgBox("同名のブックがすでに存在します。上書き保存しますか？", vbYesNo + vbExclamation) = vbNo Then
            'ここで処理を終了しているが、別名保存にするとか、方法は色々ある
            Exit Sub
            
        End If
    End If
    '---------------------------------------------------
    'ここから処理
    '---------------------------------------------------
    'ブック新規追加
    Set wb = Workbooks.Add
    Application.DisplayAlerts = False
    wb.SaveAs ThisWorkbook.Path & "\" & vFileName
    Application.DisplayAlerts = True
    wb.Close
End Sub

'---------------------------------------------------
'他のブックの内容を取得する(一度ブックを開いてそのブックを取得→閉じる)
'---------------------------------------------------
Private Sub OpenLinkedBook()
    Dim wb As Workbook
    
    '普通に開くコードだけど、ファイルの文字を書き出す部分が他ブックの参照式を含んでいると更新しますか？表示が出て止まる
    Set wb = Workbooks.Open(fileName:=ThisWorkbook.Path & "\Sample12-1.xlsx")

    MsgBox wb.Worksheets(1).Range("A1").Value
End Sub

'---------------------------------------------------
'リンクを更新せずにブックを開くコードを書いておく
'---------------------------------------------------
Private Sub OpenLinkedBook2()
    Dim wb As Workbook
    
    '普通に開くコードだけど、ファイルの文字を書き出す部分が他ブックの参照式を含んでいると更新しますか？表示が出て止まる
    'updatelinksパラメータを設定し、もしリンクが設定されていたとしてもその表示が出ないように設定しておくことが重要
    Set wb = Workbooks.Open(fileName:=ThisWorkbook.Path & "\Sample12-1.xlsx", UpdateLinks:=0)

    MsgBox wb.Worksheets(1).Range("A1").Value
End Sub

'==================================================
'下のFuncitonを呼び出す
'==================================================
Private Sub ResetSheetTest()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(1)
    
    ResetSheet sh
End Sub

'---------------------------------------------------
'すべての行と列を表示するコード
'そもそも、オートフィルターや非表示列が存在している可能性があるので、最初にそいつらを全部解除しておく
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

'仕様を最初から決めたとしても、実際の運用時には想定外の様々なケースが発生する。
'事前に想定外を想定してプログラムを作ることが大切。

'ここまでやりきってもまだゼロにならないのがエラー。
'On Error goto ステートメントのハイレベルな使い方、使うべきケースを覚える。
'==================================================

