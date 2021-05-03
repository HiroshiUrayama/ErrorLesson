VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################
'APIModuleを使用して、×ボタンを削除したUserForm
'#################################################
Option Explicit

Private Sub UserForm_Initialize()
    Dim hwnd As Long
    Dim hMenu As Long
    Dim rc As Long
    Dim vClassName As String    'クラス名
    
    'ユーザーフォームのクラス名を指定する
    vClassName = "ThunderDFrame"
    
    'ウインドウのハンドルを取得
    hwnd = FindWindow(vClassName, Me.Caption)
    
    'ウインドウに関する情報を取得
    hMenu = GetSystemMenu(hwnd, 0&)
    
    '[×]ボタンを無効にする
    rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    
    'ウインドウのメニューバーを再描画する
    rc = DrawMenuBar(hwnd)
    
End Sub

'---------------------------------------------------
'ボタンを謳歌した際の処理
'---------------------------------------------------
Private Sub CommandButton1_Click()
    Unload Me
End Sub

