Attribute VB_Name = "PyWimVBA5"
' Pywim VBA Function (5.2)
' A module to run python via VBA, By DatCanFly
' Copyright by DatCanFly (2023)
Private Function WriteToFile(filePath As String, content As String) As Boolean
    Dim fileNumber As Integer
    Dim delimiter As String
    ' Open the file for writing
    On Error GoTo ErrorHandler
    fileNumber = FreeFile
    Open filePath For Output As fileNumber
    delimiter = ";;"
    rel_cont = Split(content, delimiter)
    ' Write content to the file
    For Each item In rel_cont
        Print #fileNumber, item
    Next item
    
    
    ' Close the file
    Close fileNumber

    ' Return success
    WriteToFile = True
    Exit Function
    
ErrorHandler:
    ' An error occurred, return failure
    WriteToFile = False
End Function
Public Function RunPy(code As String, pythonPath As String, Optional outputFilePath As String = "pywvout.txt", Optional filename As String = "pywvba.py", Optional keepFileData As Boolean = False, Optional UseDebug As Boolean = False) As String
    Dim command As String
    Dim cmd As String
    Dim fileContent As String
    Dim wshShell
    If WriteToFile(filename, code) Then
        ' Nothing there
    Else
        MsgBox "[PyW.VBA] Error occurred while writing to the file."
    End If
    If UseDebug <> False Then
        command = "cmd /K " & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ "
    Else
        command = "cmd /K " & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ " & " & exit"
    End If
    'MsgBox command
    
    Set wshShell = CreateObject("WScript.Shell")
    
    If UseDebug <> False Then
        wshShell.Run command, vbNormalFocus, True
        ' # Show cmd - Debug
    Else
        wshShell.Run command, vbHide, True
    End If
    ' # Do not Show cmd - Stable
    'exitCode = wshShell.Run(command, vbHide, True)
    ' Execute the command using the Shell function
    'Shell command, vbHide
    If Dir(outputFilePath) <> "" Then
        Open outputFilePath For Input As #1
        fileContent = Input$(LOF(1), 1)
        Close #1
        If keepFileData <> True Then
            Kill outputFilePath
        End If
    Else
        fileContent = False

    End If
    If keepFileData <> True Then
        Kill filename
    End If
    RunPy = fileContent
End Function


