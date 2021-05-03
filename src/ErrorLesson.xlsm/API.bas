Attribute VB_Name = "API"
'#################################################

'#################################################
Option Explicit

'APIの命令集(32bitと64bitに対応)
'---------------------------------------------------

#If VBA7 Then
    'クラスからウインドウハンドルを取得する
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    'ウインドウに関する情報を取得する
    Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    
    'メニューから項目を削除
    Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
    'ウインドウのメニューバー外枠を再描画
    Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
 
#Else
    'クラスからウインドウハンドルを取得する
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    'ウインドウに関する情報を取得する
    Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    
    'メニューから項目を削除
    Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
    'ウインドウのメニューバー外枠を再描画
    Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

#End If


'定数の設定
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&




