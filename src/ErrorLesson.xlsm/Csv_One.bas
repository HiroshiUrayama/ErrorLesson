Attribute VB_Name = "Csv_One"
Option Explicit

'#################################################
'テキストファイルを極める
'#################################################

'==================================================
'データ行数がめちゃくちゃ多量な場合、通常のコードだと処理が重たくなるケースが多い
'==================================================

'---------------------------------------------------
'通常のcsvを読み込むファイル
'このコードは重たい(一行ずつ読み取るため)
'---------------------------------------------------
Private Sub OpenCSV()
    Dim vPath As String
    vPath = ThisWorkbook.Path & "\Data.csv"
    'csvファイルを開く
    Open vPath For Input As #1
    
    Dim i As Long
    Dim vLine As String
    
    Do Until EOF(1)
        '1行ずつ読み込む
        Line Input #1, vLine
        
        'イミディエイトウインドウに表示する
        Debug.Print vLine
    Loop
    
    'ファイルを閉じる
    Close #1
    
End Sub

'---------------------------------------------------
'CSVファイルを高速に読み込む
'これを記憶すべし
'---------------------------------------------------
Private Sub ReadCSVFile()
    Dim num As Long
    Dim buf() As Byte
    
    '空いているファイル番号を取得する
    num = FreeFile
    
    'Data.csvファイルをバイナリモードで開く
    Open ThisWorkbook.Path & "\Data.csv" For Binary As #num
    
    'ファイルの長さを取得し、変数bufの大きさを確保する
    ReDim buf(1 To LOF(num))
    Get #num, , buf     'ファイルを変数bufに読み込む
    Close #num          'ファイルを閉じる
    
    Dim DataList As Variant
    Dim temp As Variant
    Dim Data() As Variant
    Dim RowNum As Long
    Dim i As Long, j As Long
    
    '読み込んだデータを改行コードで区切り、配列に代入
    '配列は業ごとのデータになる
    DataList = Split(StrConv(buf, vbUnicode), vbCrLf)
    
    RowNum = UBound(DataList)   'データの行数を取得
    For i = 1 To RowNum                 'データの行数分処理を繰り返す
    
    '1行分のデータをカンマで区切り配列に代入
    temp = Split(DataList(i - 1), ",")
    
    '配列変数Dataの要素数を変更する
    ReDim Preserve Data(1 To RowNum, 1 To UBound(temp) + 1)
        '1行の各データを処理
        For j = 1 To UBound(temp) + 1
            'データを配列に代入
            Data(i, j) = temp(j - 1)
        Next
    Next
    
    Worksheets("Sheet2").Range("A1").Resize(UBound(Data), UBound(Data, 2)).Value = Data
    
End Sub

'---------------------------------------------------
'文字コードを変換するコード
'---------------------------------------------------

Sub StringCodeConvert()
    Call utf8ToSjis(ThisWorkbook.Path & "\Sample12-1.xlsx", ThisWorkbook.Path & "Sample12-2.xlsx")
End Sub

Public Sub utf8ToSjis(ByVal OriginPath As String, ByVal SavePath As String)
    Dim sReadData As Object
    Dim sWriteData As Object
    
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    
    Set sReadData = CreateObject("ADODB.Stream")
    Set sWriteData = CreateObject("ADODB.Stream")
    
    Dim sText As Variant
    
    sReadData.Type = adTypeText
    sReadData.Charset = "UTF-8"
    sReadData.Open
    sReadData.LoadFromFile OriginPath
    
    sWriteData.Type = adTypeText
    sWriteData.Charset = "Shift-JIS"
    sWriteData.Open
    
    sText = sReadData.readtext
    sWriteData.writetext sText
    
    sWriteData.savetofile SavePath, adSaveCreateOverWrite
    
    sReadData.Close
    sWriteData.Close
    
End Sub
