import pandas as pd
import numpy as np
import sqlite3
from sklearn.metrics import pairwise_distances
from gensim.models import KeyedVectors
import ctypes
from sqlite3 import Error

from flask import Flask, request, render_template_string, render_template, flash, redirect, url_for, session  # Using the Flask package for Web application
import math # Using the math package to calculate

# drive.google.com/file/d/0B7XkCwpI5KDYNlNUTTlSS21pQmM/edit?resourcekey=0-wjGZdNAUop6WykTtMip30g
# Load Google's pre-trained Word2Vec model.

# IMPORTANT": Terminal: conda install -c conda-forge python-levenshtein

app = Flask(__name__)
app.secret_key = "123"

con = sqlite3.connect("Database/CSDL_Thuctap.db")
con.execute("create table if not exists tbDangNhap (pid integer primary key,name text,mail text,password text)")
con.close()


# Phước làm từ đây
def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)
    return conn

def select_all_tasks(conn, my_query):
    cur = conn.cursor()
    cur.execute(my_query)

    rows = cur.fetchall()

    my_list = []
    for row in rows:
        my_list += row
    res = pd.Series(data=my_list)
    return res

conn = create_connection("Database/CSDL_Thuctap.db")
topicname_list = select_all_tasks(conn, "SELECT TenDeTaiThucTap FROM 'tbThucTap' LIMIT 0,30")
gvhdname_list = select_all_tasks(conn, "SELECT GVHD FROM 'tbThucTap' LIMIT 0,30")
# Đến đây

loaded_model = KeyedVectors.load_word2vec_format(
    "source/baomoi.model.bin", binary=True)

@app.route('/')
def user():
    return render_template('/login.html')

@app.route('/login', methods=["GET", "POST"])
def login():
    if request.method == 'POST':
        mail = request.form['mail']
        password = request.form['password']
        con = sqlite3.connect("Database/CSDL_Thuctap.db")
        con.row_factory = sqlite3.Row
        cur = con.cursor()
        cur.execute("select * from tbDangNhap where mail=? and password=?" , (mail, password))
        data = cur.fetchone()

        if data:
            if data["mail"] != "hoangtr2505@gmail.com":
                session["mail"] = data["mail"]
                session["password"] = data["password"]
                return redirect("index")
            else:
                session["mail"] = data["mail"]
                session["password"] = data["password"]
                return redirect("admin")
        else:
            flash("Mail hoặc Mật khẩu không đúng", "danger")
    return redirect(url_for("user"))


@app.route('/admin', methods=["GET", "POST"])
def admin():
    return render_template("admin.html")

@app.route('/index', methods=["GET", "POST"])
def indexuser():
    return render_template("index.html")

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        try:
            name = request.form['name']
            mail = request.form['mail']
            password = request.form['password']
            con = sqlite3.connect("Database/CSDL_Thuctap.db")
            cur = con.cursor()
            cur.execute("insert into tbDangNhap(name,mail,password)values(?,?,?)", (name, mail, password))
            con.commit()
            flash("Đăng ký tài khoản thành công", "success")
        except:
            flash("Đăng ký tài khoản thất bại!", "danger")
        finally:
            return redirect(url_for("user"))
            con.close()

    return render_template('register.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for("user"))

vocabulary = loaded_model.key_to_index.keys()
w2v_TenDeTaiThucTap = []
for i in topicname_list:
    w2Vec_word = np.zeros(400, dtype="float32")
    for word in str(i).split():
        if word in vocabulary:
            w2Vec_word = np.add(w2Vec_word, loaded_model[word])
    w2Vec_word = np.divide(w2Vec_word, len(str(i).split()))
    w2v_TenDeTaiThucTap.append(w2Vec_word)
w2v_TenDeTaiThucTap = np.array(w2v_TenDeTaiThucTap)

@app.route('/', methods=['GET', 'POST'])
def index():
    q = ""
    n = 0
    rs = ""
    if request.method == 'POST': # Receiving values through posting values from a form
        q = request.form["query"]
        n = int(request.form["num"])
        rs = f_dant_huongdan_nckhsv_2021_2022(q, n)

        rs = rs.rename(columns={'TenDeTaiThucTap': 'Tên đề tài thực tập',
                                'GVHD': 'Tên GVHD',
                                       'DoTuongTu': 'Độ tương tự'
                                       })

        return render_template('index.html', v_q=q, v_n = n, v_rs=rs.to_html(columns={
        'Độ tương tự',
        'Tên GVHD',
        'Tên đề tài thực tập'

    },
        index=False,
        justify='left'))

    else:
        return render_template('index.html', v_q='', v_n=0)


def f_dant_huongdan_nckhsv_2021_2022(query, top_number_topics):
    v_query = np.zeros(400, dtype="float32")
    for word in query.split():
        if word in vocabulary:
             v_query = np.add(v_query, loaded_model[word])
    v_query = np.divide(v_query, len(query.split()))
    couple_dist = pairwise_distances(w2v_TenDeTaiThucTap, v_query.reshape(1, -1))
    indices = np.argsort(couple_dist.ravel())[0:top_number_topics]
    df = pd.DataFrame({
                         'TenDeTaiThucTap': topicname_list[indices].values,
                         'GVHD': gvhdname_list[indices].values,
                         'DoTuongTu': couple_dist[indices].ravel(),
                     })
    print("Tìm kiếm tên đề tài thực tập","-"*50)
    print('Các đề tài liên quan đến : ', query)
    print("\n", "Kết quả tìm kiếm : ", "="*40)
    df = df.sort_values(by='DoTuongTu', ascending=True)
    df.to_html()
    return df
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
             # ctypes.windll.user32.MessageBoxW(0, "Thêm dữ liệu thành công!", "Xác nhận", 1)
             msg = "Thêm dữ liệu thành công"
        except:
            con.rollback()
            msg = "Có lỗi khi thêm dữ liệu"
        finally:
          if msg == "Thêm dữ liệu thành công":
            return render_template("SaveData.html", msg=msg)
          else:
            return render_template("SaveData2.html", msg=msg)
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
    # ctypes.windll.user32.MessageBoxW(0, "Bạn có chắc chắn muốn xóa đề tài này?", "Xoá thiệt hả", 1)
    mssv = request.args.get('mssv')
    with sqlite3.connect("Database/CSDL_Thuctap.db") as con:
     try:
            cur = con.cursor()
            cur.execute("delete from tbThuctap where MSSV =(?)", (mssv,))

            msg = "Xóa dữ liệu thành công"
     except:
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
if __name__ == '__main__':
    app.debug = True
    app.run()


'''
Đề tài nghiên cứu khoa học sinh viên năm học 2021-2022
Các thành viên:
1.
2.
3.
4.
5.
Giáo viên hướng dẫn: TS. Nguyễn Thôn Dã
'''