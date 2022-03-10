from flask import *
import sqlite3
import pandas as pd
import ctypes

from v2 import f_dant_huongdan_nckhsv_2021_2022

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("admin.html");

@app.route("/AddData")
def AddData():
    return render_template("AddData.html")

@app.route("/SaveData", methods=["POST", "GET"])
def SaveData():
    msg = "msg"
    if request.method == "POST":
        try:
            mssv = request.form["mssv"]
            ho = request.form["ho"]
            ten = request.form["ten"]
            sdt = request.form["sdt"]
            email = request.form["email"]
            tendetaithuctap = request.form["tendetaithuctap"]
            gvhd = request.form["gvhd"]
            tencongty = request.form["tencongty"]
            diachi = request.form["diachi"]
            hotennguoihuongdan = request.form["hotennguoihuongdan"]
            sdtnguoihuongdan = request.form["sdtnguoihuongdan"]
            emailnguoihuongdan = request.form["emailnguoihuongdan"]

            with sqlite3.connect("Database/CSDL_Thuctap.db") as con:
             cur = con.cursor()
             cur.execute("INSERT INTO tbThucTap (MSSV, Ho, Ten, SDT, Email, TenDeTaiThucTap, GVHD, TenCongTy, "
                         "Đia_chi, HoTen_NguoiHD, SDT_cty, Email_cty) values (?,?,?,?,?,?,?,?,?,?,?,?)",
                         (mssv, ho, ten, sdt, email, tendetaithuctap, gvhd, tencongty, diachi, hotennguoihuongdan, sdtnguoihuongdan, emailnguoihuongdan))
             con.commit()
             msg = "Thêm dữ liệu thành công"
             ctypes.windll.user32.MessageBoxW(0, "Your text", "Your title", 1)
        except:
            con.rollback()
            msg = "Có lỗi khi thêm dữ liệu"
        finally:
          return render_template("SaveData.html", msg=msg)
          con.close()



@app.route("/ShowList")
def ShowList():
    con = sqlite3.connect("Database/CSDL_Thuctap.db")
    con.row_factory = sqlite3.Row
    cur = con.cursor()
    cur.execute("select * from tbThuctap")
    rows = cur.fetchall()
    return render_template("ShowList.html", row=rows)



@app.route("/DeleteData", methods=["GET"])
def DeleteData():
    mssv = request.args.get('mssv')
    with sqlite3.connect("Database/CSDL_Thuctap.db") as con:
     try:
            cur = con.cursor()
            cur.execute("delete from tbThuctap where MSSV =(?)", (mssv,))

            ctypes.windll.user32.MessageBoxW(0, "Your text", "Your title", 1)
            msg = "Xóa dữ liệu thành công r hí"
     except:
            ctypes.windll.user32.MessageBoxW(0, "Your text", "Your title", 1)
            msg = "Không thể xóa dữ liệu"
     finally:
      return render_template("DeleteData.html", msg=msg, mssv=mssv)



@app.route('/upload', methods=['GET', 'POST'])
def upload():
    return render_template('upload.html')

@app.route('/showdata', methods=['GET','POST'])
def showdata():
        if request.method == 'POST':
            my_file = request.form['file']
            tb = pd.ExcelFile(my_file)
            tb = tb.parse(tb.sheet_names[0])
            return render_template('showdata.html', tb=tb.to_html())


if __name__ == "__main__":
    app.run(debug=True)