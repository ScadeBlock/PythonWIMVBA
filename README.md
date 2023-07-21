# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
**✨✨ New on PythonWimVBA 5.3 ✨✨**
+ Added multiple threads (Run multiple RunPy functions)
+ Improve function, add new `newengine` 
+ Enhanced PyWimVBA performance
+ Keeps the old PyWimVBA function (In version 5.2) and renames it to `RunPyOld`
- ~~Add new `RunPyWid()` - with `showcmd=True` function doesn't need to create file~~ (Removed!)
# Where's the lower version of PythonWimVBA?
> **Lower Version is tested privately; some versions are publicly released, but they are pre-release. They're outdated, unsecure and unstable, so please use only versions 5.2 or above**.

# Installation
+ Install python (Any version, NOTE: Checks Add python to path when start setup)
+ Install PyWimVBA to your vba project (Install .bas file and see [here](https://support.tetcos.com/support/solutions/articles/14000143233-how-to-import-vba-script-bas-file-in-ms-excel-) )

# Usage
**Command:** ``LoadPy(file,[iline = False])``
+ Load a python file and convert it to PyWimVBA Syntax
  - For example, it'll convert this file code
  
  `examplefile.py`

  ```
  import time
  time.sleep(5)
  ```
  - To `import time;time.sleep(5)`
  - APPLY TO USE: `RunPy(LoadPy("mycodefile.anyextension"))`
    
**Command:** ``RunPy(code,[pythonPath = "python"] , [newengine = False],  [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ [Optional] newengine : Improve performance, doesn't need to create logfile - still create code file . But it's will show up a cmd splash - doesn't work with UseDebug
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code and keep it alive with it's output [Use debug to catch errors, the output file may not catch them. So when debug is enabled, Output file does nothing.
+ E.x : `MsgBox RunPy("Welcome to \'Python With VBA!\'")`

> In lastest fix, Remove RunPyOld, RunPyWid (Issued and Too Unstable)

**Command:** `RunPyWid(code, [pythonPath = "python"], [showcmd = True], [iline = False], [UseDebug =False])`
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ With `showcmd=True` performance will be better than `showcmd=False` (because of `showcmd=False` must create logfile to catch log, `showcmd=True` mustn't)
- **[Optional OR Compulsory]iline**
  - `Iline` attributes convert code from
  - `import time`
  - `time.sleep(2)`
  - To `exec("import time\ntime.sleep(2)")`
  - (convert multiple line to single line)
+ :warning: `Iline` is custom with `showcmd=True` but it's must for `showcmd=False`
+ [Optional] UseDebug: Show-up cmd that runs python code, keep it alive with it's output [Use debug to catch errors, Output file may not catch them. Only works with `showcmd=False`]
- New release (Unstable!)
  - Aready know bugs : Syntax error (Error with `\"` and `\'` in string python, e.x: `RunPyWid("print('This function failed with \'')")` , )
    
**Command:** ``RunPyOld(code,pythonPath, [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code and keep it alive with it's output [Use debug to catch errors, the output file may not catch them. So when debug is enabled, Output file does nothing.


