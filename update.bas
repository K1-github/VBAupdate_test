Attribute VB_Name = "Main"
Public memoryConfirmed As Boolean
Public buyerID As String


Sub 作業選択画面表示()
    
    F00_作業選択.Show vbModeless
    
End Sub

'============================
'
'   メニューに戻る
'   標準モジュールからUnload
'
'============================
Sub Close_F01_Form()
    Unload F01_納品と出荷
    Set F01_納品と出荷 = Nothing
    F00_作業選択.Show
End Sub
'Sub Close_F02_Form()
'    Unload F02_同梱品の出荷
'    Set F02_同梱品の出荷 = Nothing
'    F00_作業選択.Show
'End Sub
Sub Close_F03_Form()
    Unload F03_納品_在庫化
    Set F03_納品_在庫化 = Nothing
    F00_作業選択.Show
End Sub
Sub Close_F04_Form()
    Unload F04_納品_同梱品待ち
    Set F04_納品_同梱品待ち = Nothing
    F00_作業選択.Show
End Sub


'============================
'
'   作業締め切り
'
'============================
Sub 作業締め切り()
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("本日の作業を締め切ってレポートを送信しますか？", vbOKCancel + vbExclamation, "締め切り確認")
    
    If ans = vbOK Then
        UploadFileWithMessageToChatwork
    End If

End Sub

