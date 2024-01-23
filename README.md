# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
# **🖥 PyWimVBA 6.0! - New release! ⚔**
> New features, with many changes from 5.x

# 🎉🎉**New Features**🎉🎉
- Create new RunPy
	- Require `PyServer` (1.0 for PyWVBA 6.0) - (Program will auto download it when necessary
	- Added cache value
  - Huge Performance improves from 5.x 
+ Added some server function like `ClearPyServer()` - Clear all variable cached
+ Don't need to create logfile or codefile
+ Keep 5.4 `RunPy()` and renamed it to `RunPyOld()`
+ ⚠ Warning: `PyServer 1.0` using localhost with port : `9812` 

# Installation
+ Install python (Any version, NOTE: Checks Add python to path when start setup)
+ Install PyWimVBA to your vba project (Install .bas file and see [here](https://support.tetcos.com/support/solutions/articles/14000143233-how-to-import-vba-script-bas-file-in-ms-excel-) )
### Where's the lower version of PythonWimVBA?
> **Lower Version is tested privately; some versions are publicly released, but they are pre-release. They're outdated, unsecure and unstable, so please use only versions 5.2 or above**.

# 🎨 PyWimVBA Example
### Quick Start
> Introduce basic uses of `PyWimVBA` 
+ Example construct
```
Sub MySub()
    StartPyServer ' Start the server
    ' Code here
    EndPyServer ' Stop the server
End Sub
```
FACT: You can hide the CMD of `PyWimVBA` by set `silent=true` (Read the documention below)

+ Run python code & extract python value 
```
Sub MySub()
    'Do a equation in python and export show it with MsgBox
    Dim Output as String
    StartPyServer ' Start the server
    RunPy("myvariable = 8 + 5")
    Output = RunPy("print(myvariable)") 'You can extract python value by using "print"
    MsgBox Output
    EndPyServer ' Stop the server
End Sub
```
FACT: You can hide the CMD of `PyWimVBA` by set `silent=true` (Read the documention below)

	

### Documention
#### 💎PyWimVBA 6.0 New Function
+ **`StartPyServer(Optional: pythonPath = "python",Optional useCustomPyServer: False,Optional : silent=false)`** :  Starts **`PyServer`** (PyServer supports only single server) **`silent=true`** will hide command prompt of `PyServer`. Set `useCustomPyServer` to path of your PyServer (can use in offline or a older pyserver)
+ **`EndPyServer(Optional deletePyServer: True)`** - A server doesn't close automatically with VBA, so you need to close it before end VBA code .  Use `deletePyServer=False` to keep your PyServer to reuse in `StartPyServer` (Not recommend)
+ `CheckPyServer()` - Check a or any server is running.
+ `PathPyServer()` - Get PyServer path
+ `ClearPyServer()` - Because `PyServer` support only one server, you can use this function to clear cached data like value,etc...
+ **`RunPy(code)`** - Simply run a code of python, split new line with `;;` and indentation are space,tab or tab smart syntax: `!tab~`
+ **``LoadPy(file)``** - Load a python file and convert it to PyWimVBA Syntax
  - For example, it'll convert this file code
  
  `examplefile.py`

  ```
  import time
  time.sleep(5)
  ```
  - To `import time;time.sleep(5)`
  - APPLY TO USE: `RunPy(LoadPy("mycodefile.anyextension"))`
    
+ ``RunPyOld(code,[pythonPath = "python"] , [newengine = False],  [ keepFileData = False] , [UseDebug  = False])``
> **`RunPyOld`** is "run" function from older version. I suggest you don't use it because it's outdated and buggy 
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ [Optional] newengine : Improve performance, doesn't need to create logfile - still create code file . But it's will show up a cmd splash - doesn't work with UseDebug
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code and keep it alive with it's output [Use debug to catch errors, the output file may not catch them. So when debug is enabled, Output file does nothing.
+ E.x : `MsgBox RunPy("Welcome to \'Python With VBA!\'")`

