Attribute VB_Name = "Main"
Public memoryConfirmed As Boolean
Public buyerID As String


Sub ��ƑI����ʕ\��()
    
    F00_��ƑI��.Show vbModeless
    
End Sub

'============================
'
'   ���j���[�ɖ߂�
'   �W�����W���[������Unload
'
'============================
Sub Close_F01_Form()
    Unload F01_�[�i�Əo��
    Set F01_�[�i�Əo�� = Nothing
    F00_��ƑI��.Show
End Sub
'Sub Close_F02_Form()
'    Unload F02_�����i�̏o��
'    Set F02_�����i�̏o�� = Nothing
'    F00_��ƑI��.Show
'End Sub
Sub Close_F03_Form()
    Unload F03_�[�i_�݌ɉ�
    Set F03_�[�i_�݌ɉ� = Nothing
    F00_��ƑI��.Show
End Sub
Sub Close_F04_Form()
    Unload F04_�[�i_�����i�҂�
    Set F04_�[�i_�����i�҂� = Nothing
    F00_��ƑI��.Show
End Sub


'============================
'
'   ��ƒ��ߐ؂�
'
'============================
Sub ��ƒ��ߐ؂�()
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("�{���̍�Ƃ���ߐ؂��ă��|�[�g�𑗐M���܂����H", vbOKCancel + vbExclamation, "���ߐ؂�m�F")
    
    If ans = vbOK Then
        UploadFileWithMessageToChatwork
    End If

End Sub

