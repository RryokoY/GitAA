Option Explicit

'�G���[���������Ă����f���Ȃ�
On Error Resume Next

Dim objStdIn    '�W�����͗p�I�u�W�F�N�g
Dim objStdOut   '�W���o�͗p�I�u�W�F�N�g
Dim intExitCode '�I���R�[�h

'�W�����͗p�I�u�W�F�N�g�̃C���X�^���X�쐬
Set objStdIn = Wscript.StdIn

'�W���o�͗p�I�u�W�F�N�g�̃C���X�^���X�쐬
Set objStdOut = Wscript.StdOut

'�W�����͂�1�s���ǂݍ��ݕW���o�͂֏�������

Do While objStdIn.AtEndOfStream = false

  objStdOut.WriteLine  objStdIn.ReadLine

Loop

'�C���X�^���X�̔j��
Set objStdIn   = Nothing
Set objStdOut  = Nothing

If Err.Number <> 0 Then
    intExitCode = 1 '�G���[
Else
    intExitCode = 0 '����
End If

'Quit���\�b�h�̈����l�́A�o�b�`�t�@�C����errorlevel�ɂȂ�B
Wscript.Quit(intExitCode)