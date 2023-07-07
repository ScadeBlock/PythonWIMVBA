Attribute VB_Name = "PyWimVBA5"
' PyWimVBA (5.3)
' A module to run python via VBA, By DatCanFly
' Copyright by DatCanFly (2023)

Function GenerateRandomString(length As Integer) As String
    Dim randomString As String
    Dim charSet As String
    Dim i As Integer
    
    charSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    For i = 1 To length
        randomString = randomString & Mid$(charSet, Int((Len(charSet) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = randomString
End Function
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

Function RemoveLineBreak(myString)
    If Len(myString) > 0 Then
        If Right$(myString, 2) = vbCrLf Or Right$(myString, 2) = vbNewLine Then
            myString = Left$(myString, Len(myString) - 2)
        End If
    End If
    RemoveLineBreak = myString
End Function
Function MdClean(str As String)
    str = Replace(str, """", "'")
    MdClean = str
End Function

Function Parse(stra As String)
    Dim delimiter As String
    delimiter = ";;"
    rel_cont = Split(stra, delimiter)
    Dim newstr As String
    For Each item In rel_cont
        newstr = newstr & item & vbCrLf
    Next item
    newstr = RemoveLineBreak(newstr)
    Parse = newstr
End Function
Public Function RunPy(code As String, Optional pythonPath As String = "python", Optional newengine As Boolean = False, Optional keepFileData As Boolean = False, Optional UseDebug As Boolean = False) As String
    Dim command As String
    Dim cmd As String
    Dim fileContent As String
    Dim wshShell
    Dim filename As String
    Dim outputFilePath As String
    filename = GenerateRandomString(5) & ".txt"
    outputFilePath = GenerateRandomString(6) & ".py"
    
    If WriteToFile(filename, code) Then
        ' Nothing there
    Else
        MsgBox "[PyW.VBA] Error occurred while writing to the file."
    End If
    Dim attr As String
    If UseDebug = True Then
        command = " /K "
    Else
        command = " /C "
    End If
    'MsgBox command
    If newengine = False Then
        command = "cmd" & attr & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ "
    ElseIf UseDebug = True Then
        command = "cmd" & attr & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ "
    Else
        command = """" & pythonPath & """" & " """ & filename & """ "
    End If
    Set wshShell = CreateObject("WScript.Shell")
    
    If UseDebug = True Then
        wshShell.run command, vbNormalFocus, True
        ' # Show cmd - Debug
    ElseIf newengine = False Then
        wshShell.run command, vbHide, True
    Else
        Dim gtp As Object
        Set gtp = wshShell.exec(command)

    End If
    ' # Do not Show cmd - Stable
    'exitCode = wshShell.Run(command, vbHide, True)
    ' Execute the command using the Shell function
    'Shell command, vbHide
    If newengine = False Then
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
    Else
        fileContent = gtp.stdOut.ReadAll
    End If
    If keepFileData <> True Then
        Kill filename
    End If
    RunPy = fileContent
End Function
Public Function RunPyOld(code As String, pythonPath As String, Optional keepFileData As Boolean = False, Optional UseDebug As Boolean = False) As String
    Dim command As String
    Dim cmd As String
    Dim fileContent As String
    Dim wshShell
    Dim filename As String
    Dim outputFilePath As String
    filename = GenerateRandomString(5) & ".txt"
    outputFilePath = GenerateRandomString(6) & ".py"
    
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
        wshShell.run command, vbNormalFocus, True
        ' # Show cmd - Debug
    Else
        wshShell.run command, vbHide, True
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
    RunPyOld = fileContent
End Function
Public Function oline(code)
    MsgBox run1
    Dim delimiter As String
    Dim newstr As String
    delimiter = ";;"
    nms = Split(code, delimiter)
    For Each item In nms
        newstr = newstr & item & "\n"
    Next item
    newstr = Replace(newstr, "'", "\'")
    newstr = Replace(newstr, """", "\'")
    newstr = Replace(newstr, "\\'", "\'")
    newstr = Replace(newstr, "\\" & """", "\'")
    newstr = "exec('" & newstr & "')"
    oline = newstr

End Function

Public Function RunPyWid(code As String, Optional pythonPath As String = "python", Optional showcmd As Boolean = True, Optional iline As Boolean = False, Optional UseDebug As Boolean = False)
    code = MdClean(code)
    Dim command As String
    Dim cmd As String
    Dim fileContent As String
    Dim wshShell
    If iline = False Then
        code = Parse(code)
    Else
        code = oline(code)
    End If
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim exec As Object
    If showcmd = True Then
        command = """" & pythonPath & """" & " -c """ & code & """"
        Set exec = shell.exec(command)
        Dim stdOut As String
        stdOut = exec.stdOut.ReadAll
        RunPyWid = stdOut
    Else
        Dim attr As String
        If UseDebug = False Then
            attr = " /C "
        Else
            attr = " /K "
        End If
        Dim outputFilePath As String
        outputFilePath = GenerateRandomString(6) + ".txt"
        command = "cmd" & attr & " """ & pythonPath & " -c " & """" & code & """" & " > " & outputFilePath & """"
        If UseDebug = False Then
            shell.run command, vbHide, True
        Else
            shell.run command, vbNormalFocus, True
        End If

        If Dir(outputFilePath) <> "" Then
            Open outputFilePath For Input As #1
            RunPy = Input$(LOF(1), 1)
            Close #1
            Kill outputFilePath
        Else
            RunPyWid = False
        End If
    End If
End Function
Public Function LoadPy(file, Optional iline As Boolean = False)
    Dim code As String
    Open file For Input As #1
    code = Input$(LOF(1), 1)
    Close #1
    'code = Replace(code, vbCr, ";;")
    code = Replace(code, vbCrLf, ";;")
    code = Replace(code, vbLf, ";;")
    code = Replace(code, vbLf, ";;")
    code = Replace(code, vbLf, ";;")

    If iline = True Then
        code = oline(code)
        
    End If
    LoadPy = code
End Function

Sub mypycode()
    Dim code As String
    code = LoadPy("code.txt")
    MsgBox RunPy("print('Welcome to \'Python With VBA!\'')", , True)
End Sub
