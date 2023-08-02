# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
# **🖥 PyWimVBA 6.0! - New release! ⚔**
> New features, with many changes from 5.x
### 🎉🎉**New Features**🎉🎉
- Create new RunPy
	- Require `PyServer` (1.0 for PyWVBA 6.0) 
	- Added cache value
  - Huge Performance improves from 5.x 
+ Added some server function like `ClearPyServer()` - Clear all variable cached
+ Don't need to create logfile or codefile
+ Keep 5.4 `RunPy()` and renamed it to `RunPyOld()`
+ ⚠ Warning: `PyServer 1.0` using localhost with port : `9812` 
# Where's the lower version of PythonWimVBA?
> **Lower Version is tested privately; some versions are publicly released, but they are pre-release. They're outdated, unsecure and unstable, so please use only versions 5.2 or above**.

# Installation
+ Install python (Any version, NOTE: Checks Add python to path when start setup)
+ Install PyWimVBA to your vba project (Install .bas file and see [here](https://support.tetcos.com/support/solutions/articles/14000143233-how-to-import-vba-script-bas-file-in-ms-excel-) )

# 🎨 PyWimVBA Example
### Documention
#### 💎PyWimVBA 6.0 New Function
+ `StartPyServer(Optional: pythonPath = "python",Optional useCustomPyServer: False,Optional : silent=false)` :  Starts **`PyServer`** (PyServer supports only single server) `silent=true` will hide command prompt of `PyServer`. Set `useCustomPyServer` to path of your PyServer (can use in offline or a older pyserver)
+ `EndPyServer(Optional deletePyServer: True)` - A server doesn't close automatically with VBA, so you need to close it before end VBA code .  Use `deletePyServer=False` to keep your PyServer to reuse in `StartPyServer` (Not recommend)
+ `CheckPyServer()` - Check a or any server is running.
+ `ClearPyServer()` - Because `PyServer` support only one server, you can use this function to clear cached data like value,etc...
+ `RunPy(code)` - Simply run a code of python, split new line with `;;`
+ ``LoadPy(file)`` - Load a python file and convert it to PyWimVBA Syntax
  - For example, it'll convert this file code
  
  `examplefile.py`

  ```
  import time
  time.sleep(5)
  ```
  - To `import time;time.sleep(5)`
  - APPLY TO USE: `RunPy(LoadPy("mycodefile.anyextension"))`
    
+ ``RunPyOld(code,[pythonPath = "python"] , [newengine = False],  [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ [Optional] newengine : Improve performance, doesn't need to create logfile - still create code file . But it's will show up a cmd splash - doesn't work with UseDebug
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code and keep it alive with it's output [Use debug to catch errors, the output file may not catch them. So when debug is enabled, Output file does nothing.
+ E.x : `MsgBox RunPy("Welcome to \'Python With VBA!\'")`
### Example usage
#### Here's a example code to run python in vba wuth PWV
```
Sub running()
    StartPyServer
    RunPy ("example_value = 'Hello from PWA 6!'")
    MsgBox RunPy("print(example_value)") 'To test cached value
    MsgBox RunPy("if 1+1==2:;;  print('It actually works!')")
    EndPyServer
End Sub
```
