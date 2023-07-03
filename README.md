# PythonWIMVBA
> Python With VBA ! Run python via VBA
---
**✨✨ New on PythonWimVBA 5.2 ✨✨**
+ Catch All Python Output, so outtl is useless (removed!, don't need on 5.2 or higher version)
+ Added UseDebug (show cmd and keep cmd alive with it's output) and keepFileData (keeps pywvout.txt and pywvba.py after execution) attributes

# Where're lower version of PythonWimVBA?
**Lower Version is tested privately - some version is public released but they are pre-release. They're outdated, unsecure and unstable, so please use only version 5.2 or above**

# Usage
**Command:** ``RunPy(code,pythonPath, [outputFilePath = "pywvout.txt"] , [ filename = "pywvba.py"], [ keepFileData = False] , [UseDebug  = False])``
+ Code splitting by ";;" , e.x : "import time;;time.sleep(5)"
+ [Optional] Output File Path - File that write every output of python
