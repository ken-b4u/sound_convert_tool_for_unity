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
        inputValue = InputBox("bitrate 64-256", "Input")
        if inputValue >= 64 and inputValue <= 256 Then
            For Each objItem In fileList.Files
                strSplit = Split(objItem.Name, ".")
                ExecuteEncode objItem.Name, strSplit(0) & ".mp3", inputValue
                'ExecuteEncode objItem.Name, strSplit(0) & ".m4a", inputValue
            Next
        End if
        Set objFileSys = Nothing
        Set objShell = Nothing
        Set strScriptPath = Nothing
        Set strBinPath = Nothing
        MsgBox "Finished"
    End Function
    Private Sub ExecuteEncode(inputFileName, outputFileName, inputValue)
        Set objShell = CreateObject("WScript.Shell")
        'strCommand = strBinPath & " -i " & inputDirectory & "\" & inputFileName & " -c:a libfdk_aac -b:a " & inputValue & "000 " & outputDirectory & "\" & outputFileName
        strCommand = strBinPath & " -i " & inputDirectory & "\" & inputFileName & " -b:a " & inputValue & "000 " & outputDirectory & "\" & outputFileName
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
