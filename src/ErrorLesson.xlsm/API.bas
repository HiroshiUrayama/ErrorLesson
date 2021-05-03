Attribute VB_Name = "API"
'#################################################

'#################################################
Option Explicit

'API�̖��ߏW(32bit��64bit�ɑΉ�)
'---------------------------------------------------

#If VBA7 Then
    '�N���X����E�C���h�E�n���h�����擾����
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    '�E�C���h�E�Ɋւ�������擾����
    Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    
    '���j���[���獀�ڂ��폜
    Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
    '�E�C���h�E�̃��j���[�o�[�O�g���ĕ`��
    Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
 
#Else
    '�N���X����E�C���h�E�n���h�����擾����
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    '�E�C���h�E�Ɋւ�������擾����
    Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    
    '���j���[���獀�ڂ��폜
    Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
    '�E�C���h�E�̃��j���[�o�[�O�g���ĕ`��
    Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

#End If


'�萔�̐ݒ�
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&




