Option Explicit

'���̃X�N���v�g�ȉ��̃t�@�C�����������āAimg[i] = "filepath";�̌`�ɂ���X�N���v�g�ł��B
'�g�D�摜��C�ӂ̃t�H���_�ɓ���A�g�D���~�{��.�g���q�̃t�H�[�}�b�g�Ŗ��Â��Ă�������
'�~�����݂��邩�ۂ��ł��̉摜��K�v�ȉ摜�����肵�Ă��܂��B
'�������ꂽ�e�L�X�g���Ahtml��Javascript�����ɓ\��t����Ƃ��������ɂȂ�܂��B
'DelStr�ɂ͂��̃X�N���v�g������Ƃ��܂ł̃t���p�X������āB
'Made by 314�̐l

Dim FIND_START_FOLDER
FIND_START_FOLDER = ".\"                    '�T���J�nfolder
Dim FIND_RESULT_FILE_NAME
FIND_RESULT_FILE_NAME = ".\FIND_RESULT.TXT" '�T�����ʈꗗ
Dim FIND_RESULT_FILE_OBJ
Dim i
i = 0
Dim objRE
Dim objMatches
Set objRE = CreateObject("VBScript.RegExp")
Dim DelStr

DelStr = "" '���ʂȕ���

Sub Main()
    
    Dim objFSO          ' FileSystemObject
    
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    'refer to http://msdn.microsoft.com/ja-jp/library/ie/cc428044.aspx
    '2=�����p�Ƃ���open  , True=file�V�K�쐬 , -1=unicode�ŏ���
    Set FIND_RESULT_FILE_OBJ = objFSO.OpenTextFile(FIND_RESULT_FILE_NAME, 2, True, -1)    
    FindFolder objFSO.getFolder(FIND_START_FOLDER)
    
    FIND_RESULT_FILE_OBJ.Close
    
End Sub


' �t�H���_�����֐�
Sub FindFolder(ByVal objParentFolder)
    
    Dim objFile
    Dim resultLine
    Dim FileName
    Dim FileNameTmp
    
    For Each objFile In objParentFolder.Files
        FileNameTmp = objFile.ParentFolder & "\" & objFile.Name
        If InStr(FileNameTmp, "�y") > 0 Then
            FileName = Replace(Replace(FileNameTmp, DelStr, "."), "\", "/")
            FIND_RESULT_FILE_OBJ.Write ("img[" & i & "] = """ & FileName & """;")
            FIND_RESULT_FILE_OBJ.WriteLine ("")
            i = i + 1          
        End If
    Next
    
    Dim objSubFolder    ' �T�u�t�H���_
    For Each objSubFolder In objParentFolder.SubFolders
        FindFolder (objSubFolder)
    Next
    
    Set objRE = Nothing
    
End Sub

Main
