Option Explicit

'このスクリプト以下のファイルを検索して、img[i] = "filepath";の形にするスクリプトです。
'組織画像を任意のフォルダに入れ、組織名×倍率.拡張子のフォーマットで名づけてください
'×が存在するか否かでその画像を必要な画像か判定しています。
'生成されたテキストを、htmlのJavascript部分に貼り付けるといい感じになります。
'DelStrにはこのスクリプトがあるとこまでのフルパスをいれて。
'Made by 314の人

Dim FIND_START_FOLDER
FIND_START_FOLDER = ".\"                    '探索開始folder
Dim FIND_RESULT_FILE_NAME
FIND_RESULT_FILE_NAME = ".\FIND_RESULT.TXT" '探索結果一覧
Dim FIND_RESULT_FILE_OBJ
Dim i
i = 0
Dim objRE
Dim objMatches
Set objRE = CreateObject("VBScript.RegExp")
Dim DelStr

DelStr = "" '無駄な部分

Sub Main()
    
    Dim objFSO          ' FileSystemObject
    
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    'refer to http://msdn.microsoft.com/ja-jp/library/ie/cc428044.aspx
    '2=書込用としてopen  , True=file新規作成 , -1=unicodeで書込
    Set FIND_RESULT_FILE_OBJ = objFSO.OpenTextFile(FIND_RESULT_FILE_NAME, 2, True, -1)    
    FindFolder objFSO.getFolder(FIND_START_FOLDER)
    
    FIND_RESULT_FILE_OBJ.Close
    
End Sub


' フォルダ検索関数
Sub FindFolder(ByVal objParentFolder)
    
    Dim objFile
    Dim resultLine
    Dim FileName
    Dim FileNameTmp
    
    For Each objFile In objParentFolder.Files
        FileNameTmp = objFile.ParentFolder & "\" & objFile.Name
        If InStr(FileNameTmp, "【") > 0 Then
            FileName = Replace(Replace(FileNameTmp, DelStr, "."), "\", "/")
            FIND_RESULT_FILE_OBJ.Write ("img[" & i & "] = """ & FileName & """;")
            FIND_RESULT_FILE_OBJ.WriteLine ("")
            i = i + 1          
        End If
    Next
    
    Dim objSubFolder    ' サブフォルダ
    For Each objSubFolder In objParentFolder.SubFolders
        FindFolder (objSubFolder)
    Next
    
    Set objRE = Nothing
    
End Sub

Main
