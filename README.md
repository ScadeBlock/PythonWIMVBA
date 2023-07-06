# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
**✨✨ New on PythonWimVBA 5.3 ✨✨**
+ Added multiple threads (Run multiple RunPy functions)
+ Improve function; the new `RunPy()` - with `showcmd=True` function doesn't need to create file, with `showcmd=`
+ Remove `keepFileData` attributes and add `showcmd`
+ Enhanced PyWimVBA performance
+ Keeps the old PyWimVBA function (In version 5.2) and renames it to `RunPyOld`
+ Added iline attributes
# Where's the lower version of PythonWimVBA?
> **Lower Version is tested privately; some versions are publicly released, but they are pre-release. They're outdated, unsecure and unstable, so please use only version 5.2 or above**.

# Usage
**Command:** `RunPy(code, [pythonPath = "python"], [showcmd = True], [iline = False], [UseDebug =False])`
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

**Command:** ``RunPy(code,pythonPath, [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : `import time;;time.sleep(5)`
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code and keep it alive with it's output [Use debug to catch errors, the output file may not catch them. So when debug is enabled, Output file does nothing.


---

| Information | `RunPy()` | `RunPyOld()` |
| ----------- | -----------                 | ----------- |
| Performance | :zap: (amazing) | :+1: (good) |
| Easy-Debug | :star: (very good, with `showcmd = False & UseDebug=True`,`showcmd = True` is not recommend for debugging)   |  :star2: (amazing,with `UseDebug & keepFileData=True`)|
| Easy-To-Use | :ok_hand: |:ok_hand: |
| Cache file | :raised_hands: (no cache, but `showcmd=False` needs create logfile) | :turtle: (must create logfile,code file) | 


# Some advices?
+ Although both PyWimVba Function - New and old have it debug method, i suggest you write and debug your code in a code editor like `Visual Studio Code`, `Sublime Text`. And then , replace all the newline with `;;` and paste it into your vba scripts
