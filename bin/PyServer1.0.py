# PyServer 1.0 - PyWimVBA 6.0
from flask import Flask,request
app = Flask(__name__)
from io import StringIO
from contextlib import redirect_stdout
from traceback import format_exc
import sys,os


@app.route("/")
def main():
    cdef = request.args.get("code")
    if cdef == None:
        return "PyWimVBA server, this does nothing"
    elif cdef == "$clear":
        for n in dir():
            if n[0]!='_': delattr(sys.modules[__name__], n)
        return "Cleared"
    elif cdef == "$exit":
        own_pid = os.getpid()
        os.kill(own_pid, 9)
    elif cdef == "$path":
        return __file__
    else:
        cpde = cdef
        f = StringIO()
        cpde = cpde.replace(";;","\n")
        cpde = cpde.replace("!plus~","+")
        cpde = cpde.replace("!and~","&")
        cpde = cpde.replace("!equal~","=")
        with redirect_stdout(f):
            try:
                exec(cpde, globals())
            except:
                te = format_exc()
                return te.replace("\n",";;")
        s = f.getvalue()
        return s
if __name__ == "__main__":
    app.run(debug=True,host="0.0.0.0",port="9812")
