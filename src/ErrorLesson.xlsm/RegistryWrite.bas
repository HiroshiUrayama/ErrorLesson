Attribute VB_Name = "RegistryWrite"
'#################################################
'WindowsAPI�̃}�N��
'WindowsAPI-Windows�̋@�\�𒼐�VBA���̃v���O�������ꂩ�爵�����߂̖��߁B
'VBA_�G���[���N�����Ƃ��Ă��AExcel��VBE���G���[�̔��������m���āAExcel������Ƃ��������Ƃ������悤�ɂ��Ă�B
    '�u���v����
'API�͎�鏈�������Ȃ��̂ŁA�����Ȃ�Windows���̂��u������v���Ƃ����肦��c��鏈�����Ȃ��B
'�����W�X�g���AAPI���g���̂͊댯�ƌ�����B
'#################################################
Option Explicit

'==================================================
'�Z�b�g�A�b�v���Ƀ��W�X�g���ɏ����������ރv���O����
'==================================================
Private Sub RegistrySample()
    '���W�X�g���ɃL�[��ǉ�����
    SaveSetting "VBASample", "Main", "test", "Sample"
    
    '���W�X�g������l��ǂݍ���
    MsgBox GetSetting("VBASample", "Main", "Test", "Sample")
    
End Sub

'==================================================
'���W�X�g������f�[�^���폜����R�[�h
'==================================================
Private Sub RegistryDeleteSample()
    '���W�X�g���ɃL�[��ǉ�����
    
    '���W�X�g���̃Z�N�V�������폜����
    DeleteSetting "VBASample"
End Sub

'�⑫
