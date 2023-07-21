# PyServer 1.0 - PyWimVBA 6.0
from flask import Flask,request
app = Flask(__name__)
from io import StringIO
from contextlib import redirect_stdout
from traceback import format_exc
import sys
value = {}
@app.route("/")
def main():
    if request.args.get("code") == None:
        return "PyWimVBA server, this does nothing"
    elif request.args.get("code") == "$clear":
        sys.modules[__name__].__dict__.clear()
        return "Cleared"
    else:
        cpde = request.args.get("code")
        f = StringIO()
        cpde = cpde.replace(";;","\n")
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
