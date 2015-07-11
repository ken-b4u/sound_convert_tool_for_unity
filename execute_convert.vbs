Const inputDirectory = "sound\input"
Const outputDirectory = "sound\output"
Class ExecuteConvert
    Dim strScriptPath
    Dim strBinPath
    Dim objFileSys
    Dim objShell
    Public Function Encode()
        Set objFileSys = CreateObject("Scripting.FileSystemObject")
        strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
        strBinPath = objFileSys.BuildPath(strScriptPath, "bin\ffmpeg.exe")
        Set fileList = GetFileList()
        For Each objItem In fileList.Files
            strSplit = Split(objItem.Name, ".")
            ExecuteEncode objItem.Name, strSplit(0) & ".mp3"
        Next
        MsgBox "Finished"
    End Function
    Private Sub ExecuteEncode(inputFileName, outputFileName)
        Set objShell = CreateObject("WScript.Shell")
        strCommand = strBinPath & " -i " & inputDirectory & "\" & inputFileName & " -b:v 192000 " & outputDirectory & "\" & outputFileName
        msgbox strCommand
        objShell.Run strCommand,0,False
    End Sub
    Private Function GetFileList()
        strTargetPath = objFileSys.BuildPath(strScriptPath, inputDirectory)
        Set objFolder = objFileSys.GetFolder(strTargetPath)
        Set GetFileList = objFolder
    End Function
End Class
Set objExecuteConvert = New ExecuteConvert
objExecuteConvert.Encode()
Set objExecuteConvert = Nothing
