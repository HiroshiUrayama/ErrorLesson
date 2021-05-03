Attribute VB_Name = "Csv_One"
Option Explicit

'#################################################
'�e�L�X�g�t�@�C�����ɂ߂�
'#################################################

'==================================================
'�f�[�^�s�����߂��Ⴍ���ᑽ�ʂȏꍇ�A�ʏ�̃R�[�h���Ə������d�����Ȃ�P�[�X������
'==================================================

'---------------------------------------------------
'�ʏ��csv��ǂݍ��ރt�@�C��
'���̃R�[�h�͏d����(��s���ǂݎ�邽��)
'---------------------------------------------------
Private Sub OpenCSV()
    Dim vPath As String
    vPath = ThisWorkbook.Path & "\Data.csv"
    'csv�t�@�C�����J��
    Open vPath For Input As #1
    
    Dim i As Long
    Dim vLine As String
    
    Do Until EOF(1)
        '1�s���ǂݍ���
        Line Input #1, vLine
        
        '�C�~�f�B�G�C�g�E�C���h�E�ɕ\������
        Debug.Print vLine
    Loop
    
    '�t�@�C�������
    Close #1
    
End Sub

'---------------------------------------------------
'CSV�t�@�C���������ɓǂݍ���
'������L�����ׂ�
'---------------------------------------------------
Private Sub ReadCSVFile()
    Dim num As Long
    Dim buf() As Byte
    
    '�󂢂Ă���t�@�C���ԍ����擾����
    num = FreeFile
    
    'Data.csv�t�@�C�����o�C�i�����[�h�ŊJ��
    Open ThisWorkbook.Path & "\Data.csv" For Binary As #num
    
    '�t�@�C���̒������擾���A�ϐ�buf�̑傫�����m�ۂ���
    ReDim buf(1 To LOF(num))
    Get #num, , buf     '�t�@�C����ϐ�buf�ɓǂݍ���
    Close #num          '�t�@�C�������
    
    Dim DataList As Variant
    Dim temp As Variant
    Dim Data() As Variant
    Dim RowNum As Long
    Dim i As Long, j As Long
    
    '�ǂݍ��񂾃f�[�^�����s�R�[�h�ŋ�؂�A�z��ɑ��
    '�z��͋Ƃ��Ƃ̃f�[�^�ɂȂ�
    DataList = Split(StrConv(buf, vbUnicode), vbCrLf)
    
    RowNum = UBound(DataList)   '�f�[�^�̍s�����擾
    For i = 1 To RowNum                 '�f�[�^�̍s�����������J��Ԃ�
    
    '1�s���̃f�[�^���J���}�ŋ�؂�z��ɑ��
    temp = Split(DataList(i - 1), ",")
    
    '�z��ϐ�Data�̗v�f����ύX����
    ReDim Preserve Data(1 To RowNum, 1 To UBound(temp) + 1)
        '1�s�̊e�f�[�^������
        For j = 1 To UBound(temp) + 1
            '�f�[�^��z��ɑ��
            Data(i, j) = temp(j - 1)
        Next
    Next
    
    Worksheets("Sheet2").Range("A1").Resize(UBound(Data), UBound(Data, 2)).Value = Data
    
End Sub

'---------------------------------------------------
'�����R�[�h��ϊ�����R�[�h
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
