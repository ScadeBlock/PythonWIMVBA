Attribute VB_Name = "PyWimVBA6"
' PyWimVBA (6.0) Big Jump - Beta
' A module to run python via VBA, By DatCanFly
' Copyright by DatCanFly (2023)
Private Function GenerateRandomString(length As Integer) As String
    Dim randomString As String
    Dim charSet As String
    Dim i As Integer
    
    charSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    For i = 1 To length
        randomString = randomString & Mid$(charSet, Int((Len(charSet) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = randomString
End Function
Private Function SmartSyntax(code As String)
    'code = Replace(code, vbCr, ";;")
    code = Replace(code, "!tab~", "    ")
    code = Replace(code, "+", "!plus~")
    code = Replace(code, "&", "!and~")
    code = Replace(code, "=", "!equal~")
    code = Replace(code, vbLf, ";;")
    SmartSyntax = code
End Function
Private Function WriteToFile(filePath As String, content As String, Optional spliton As Boolean = True) As Boolean
    Dim fileNumber As Integer
    Dim delimiter As String
    ' Open the file for writing
    On Error GoTo ErrorHandler
    fileNumber = FreeFile
    Open filePath For Output As fileNumber
    If spliton = True Then
        delimiter = ";;"
        rel_cont = Split(content, delimiter)
        ' Write content to the file
    Else
        delimiter = vbCrLf
        rel_cont = Split(content, delimiter)
    End If
    For Each Item In rel_cont
        Print #fileNumber, Item
    Next Item
    

    ' Close the file
    Close fileNumber

    ' Return success
    WriteToFile = True
    Exit Function
    
ErrorHandler:
    ' An error occurred, return failure
    WriteToFile = False
End Function

Private Function RemoveLineBreak(myString)
    If Len(myString) > 0 Then
        If Right$(myString, 2) = vbCrLf Or Right$(myString, 2) = vbNewLine Then
            myString = Left$(myString, Len(myString) - 2)
        End If
    End If
    RemoveLineBreak = myString
End Function

Private Function Parse(stra As String)
    Dim delimiter As String
    delimiter = ";;"
    rel_cont = Split(stra, delimiter)
    Dim newstr As String
    For Each Item In rel_cont
        newstr = newstr & Item & vbCrLf
    Next Item
    newstr = RemoveLineBreak(newstr)
    Parse = newstr
End Function
Public Function StartPyServer(Optional pythonPath As String = "python", Optional useCustomPyServer = False, Optional silent As Boolean = False)
    Dim pfo As String
    If Not useCustomPyServer = False Then
        pfo = useCustomPyServer
    Else
        pfo = Environ("Temp") & "\" & GenerateRandomString(8) & ".py"
        Dim Request As Object
        Set Request = CreateObject("MSXML2.XMLHTTP")
        
        With Request
            .Open "GET", "https://raw.githubusercontent.com/ScadeBlock/PythonWIMVBA/main/bin/PyServer1.0.py", False
            .Send
            
            If WriteToFile(pfo, .responseText, False) Then
                ' Nothing there
            Else
                MsgBox "[PyW.VBA] Error occurred while writing to the file."
            End If
        End With
        Set Request = Nothing
    End If

    
    Dim command As String
    Dim attr As String
    If silent = True Then
        attr = vbHide
    Else
        attr = vbNormalFocus
    End If
    command = "cmd /C " & " """ & """" & pythonPath & """" & " """ & pfo & """" & """ "
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run command, attr, False
End Function
Public Function CheckPyServer() As Boolean
    Dim Request As Object
    Dim ff As Integer
    Dim rc As Variant
    Dim url As String
    url = "http://127.0.0.1:9812"
    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
      .Open "GET", url, False
      .Send
      rc = .StatusText
    End With
    Set Request = Nothing
    If rc = "OK" Then CheckPyServer = True

    Exit Function
EndNow:
End Function
Public Function PathPyServer()
    With Request
        .Open "GET", "http://127.0.0.1:9812/?code=$path", False
        .Send
        PathPyServer = .responseText
    End With
End Function
Public Function EndPyServer(Optional deletePyServer = True)
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    Dim passy As String
    
        
    With Request
        .Open "GET", "http://127.0.0.1:9812/?code=$path", False
        .Send
        passy = .responseText
    End With
    Set Request = Nothing
    Set Request = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo Oops
    With Request
        .Open "GET", "http://127.0.0.1:9812/?code=$exit", False
        .Send
    End With
Oops:
    'ex
        If deletePyServer = True & Dir(passy) Then
            Kill passy
        End If

End Function
Public Function ClearPyServer()
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    Dim passy As String

    With Request
        .Open "GET", "http://127.0.0.1:9812/?code=$clear", False
        .Send
    End With
End Function
Public Function RunPy(code As String)
    code = SmartSyntax(code)
    Dim Request As Object
    Dim url As String
    url = "http://127.0.0.1:9812?code=" & code
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", url, False
    objHTTP.Send ""
    RunPy = objHTTP.responseText
End Function
Public Function RunPyOld(code As String, Optional pythonPath As String = "python", Optional newengine As Boolean = False, Optional keepFileData As Boolean = False, Optional UseDebug As Boolean = False) As String
    Dim command As String
    Dim cmd As String
    Dim fileContent As String
    Dim WshShell
    Dim filename As String
    Dim outputFilePath As String
    filename = Environ("Temp") & "\" & GenerateRandomString(5) & ".py"
    outputFilePath = Environ("Temp") & "\" & GenerateRandomString(6) & ".txt"
    
    If WriteToFile(filename, code) Then
        ' Nothing there
    Else
        MsgBox "[PyW.VBA] Error occurred while writing to the file."
    End If
    Dim attr As String
    If UseDebug = True Then
        attr = " /K "
    Else
        attr = " /C "
    End If
    'MsgBox command
    If newengine = False Then
        command = "cmd" & attr & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ "
    ElseIf UseDebug = True Then
        command = "cmd" & attr & " """ & """" & pythonPath & """" & " """ & filename & """ > """ & outputFilePath & """" & """ "
    Else
        command = """" & pythonPath & """" & " """ & filename & """ "
    End If
    Set WshShell = CreateObject("WScript.Shell")
    MsgBox command
    If UseDebug = True Then
        WshShell.Run command, vbNormalFocus, True
        ' # Show cmd - Debug
    ElseIf newengine = False Then
        WshShell.Run command, vbHide, True
    Else
        Dim gtp As Object
        Set gtp = WshShell.exec(command)
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
    RunPyOld = fileContent
End Function
Public Function LoadPy(file)
    Dim code As String
    Open file For Input As #1
    code = Input$(LOF(1), 1)
    Close #1
    'code = Replace(code, vbCr, ";;")
    code = Replace(code, vbCrLf, ";;")
    code = Replace(code, vbLf, ";;")
    code = Replace(code, vbLf, ";;")
    code = Replace(code, vbLf, ";;")
    LoadPy = code
End Function
Sub running()
    'StartPyServer
    RunPy ("example_value = 'Hello from PWA 6!'")
    MsgBox RunPy("print(example_value)") 'To test cached value
    MsgBox RunPy("if 1+1==2:;;!tab~print('It actually works!')")
    EndPyServer
End Sub
