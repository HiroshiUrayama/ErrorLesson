VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################
'APIModule���g�p���āA�~�{�^�����폜����UserForm
'#################################################
Option Explicit

Private Sub UserForm_Initialize()
    Dim hwnd As Long
    Dim hMenu As Long
    Dim rc As Long
    Dim vClassName As String    '�N���X��
    
    '���[�U�[�t�H�[���̃N���X�����w�肷��
    vClassName = "ThunderDFrame"
    
    '�E�C���h�E�̃n���h�����擾
    hwnd = FindWindow(vClassName, Me.Caption)
    
    '�E�C���h�E�Ɋւ�������擾
    hMenu = GetSystemMenu(hwnd, 0&)
    
    '[�~]�{�^���𖳌��ɂ���
    rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    
    '�E�C���h�E�̃��j���[�o�[���ĕ`�悷��
    rc = DrawMenuBar(hwnd)
    
End Sub

'---------------------------------------------------
'�{�^����搉̂����ۂ̏���
'---------------------------------------------------
Private Sub CommandButton1_Click()
    Unload Me
End Sub

