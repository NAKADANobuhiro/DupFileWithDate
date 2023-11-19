'�I�v�V�����w��
Option Explicit 

'�ϐ��錾
Dim fsObj
Dim strFileName
Dim strNewFileName
Dim strParentFolderName
Dim strBaseName
Dim ext
dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Dim regEx, strDate

'������1�łȂ���ΏI��
If WScript.Arguments.Count <> 1 Then WScript.Quit

strFileName = WScript.Arguments(0) '�t�@�C�����擾
Set fsObj = CreateObject("Scripting.FileSystemObject") '�t�@�C���V�X�e���I�u�W�F�N�g���擾
strBaseName = fsObj.GetBaseName(strFileName) '�t�@�C���̃x�[�X���擾

Set regEx = New RegExp
regEx.Global		= True	
regEx.IgnoreCase	= True	
regEx.Global = False        

regEx.pattern = "[0-9]{4}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])" '���t�炵�� yyyymmdd ���l�Ƀ}�b�`
strDate = Replace(Left(Now(),10), "/", "") '�����̓��t�� yyyymmdd �^�Ŏ擾

If regEx.Test(strBaseName) Then '�t�@�C�����ɓ��t(yyyymmdd)���܂܂�Ă����
    strBaseName = regEx.Replace(strBaseName, strDate) '���̓��t�������̓��t�ɒu��������
else
    regEx.pattern = "[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[12][0-9]|3[01])" '���t�炵�����l yyyy-mm-dd �Ƀ}�b�`
    strDate = Replace(Left(Now(),10), "/", "-") '�����̓��t�� yyyy-mm-dd �^�Ŏ擾
    If regEx.Test(strBaseName) Then '�t�@�C�����ɓ��t(yyyy-mm-dd)���܂܂�Ă����
        strBaseName = regEx.Replace(strBaseName, strDate) '���̓��t�������̓��t�ɒu��������
    else    
        strBaseName = strBaseName & "_" & strDate '�܂܂�Ă��Ȃ���΍����̓��t��t������
    end if
end if

ext = LCase(fsObj.GetExtensionName(strFileName))'�t�@�C���̊g���q�擾
strParentFolderName = fsObj.GetParentFolderName(strFileName)'�t�@�C���̐e�t�H���_���擾
strNewFileName = strParentFolderName & "\" & strBaseName & "." & ext'���t������t�^�����t�@�C�������쐬

'�R�s�[����
if strFileName <> strNewFileName then
    fsObj.CopyFile strFileName, strNewFileName, true
else
    objShell.Popup "���ɍ����̓��t�����Ă��܂��B", 5,, vbInformation
end if

'�I������
set fsObj = Nothing
set strFileName = Nothing
set strNewFileName = Nothing
set strParentFolderName = Nothing
set strBaseName = Nothing
set ext = Nothing
WScript.Quit