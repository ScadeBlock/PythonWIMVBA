# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
**✨✨ New on PythonWimVBA 5.3 ✨✨**
+ Added multiple threads (Run multiple RunPy function)
+ Remove output file Path and filename (remove!, Added auto random generate name)

# Where's the lower version of PythonWimVBA?
> **Lower Version is tested privately; some versions are publicly released, but they are pre-release. They're outdated, unsecure and unstable, so please use only version 5.2 or above**

# Usage
**Command:** ``RunPy(code,pythonPath, [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : "import time;;time.sleep(5)"
+ [Optional] keepFileData: Keep the output file and code file after finishing execution. 
+ [Optional] UseDebug: Show cmd that runs python code, keep it alive with it's output [Use debug to catch errors, Output file may not catch them. So when debug is enabled, Output file does nothing.]
