Attribute VB_Name = "RegistryWrite"
'#################################################
'WindowsAPIのマクロ
'WindowsAPI-Windowsの機能を直接VBA等のプログラム言語から扱うための命令。
'VBA_エラーが起きたとしても、ExcelやVBEがエラーの発生を検知して、Excelが壊れるといったことが無いようにしてる。
    '「守る」処理
'APIは守る処理をしないので、いきなりWindows自体が「落ちる」ことがありえる…守る処理がない。
'∴レジストリ、APIを使うのは危険と言われる。
'#################################################
Option Explicit

'==================================================
'セットアップ時にレジストリに情報を書き込むプログラム
'==================================================
Private Sub RegistrySample()
    'レジストリにキーを追加する
    SaveSetting "VBASample", "Main", "test", "Sample"
    
    'レジストリから値を読み込む
    MsgBox GetSetting("VBASample", "Main", "Test", "Sample")
    
End Sub

'==================================================
'レジストリからデータを削除するコード
'==================================================
Private Sub RegistryDeleteSample()
    'レジストリにキーを追加する
    
    'レジストリのセクションを削除する
    DeleteSetting "VBASample"
End Sub

'補足
