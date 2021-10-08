from flask import (Flask, render_template, request, session,
                   redirect, url_for, flash, send_from_directory)
import os
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from excelpy import excelpy


# Flaskをインスタンス化してsecret_keyをセットする
app = Flask(__name__)
key = os.urandom(13)
app.secret_key = key


# ログイン時に使用するIDとPW
ID_PW = {'FLASK': 'EXCEL'}


# sqliteの準備
URI = 'sqlite:///file.db'
app.config['SQLALCHEMY_DATABASE_URI'] = URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)


# ORMでDBを操作するためにクラスを継承
class DB(db.Model):
    __tablename__ = 'file_table'
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(30), unique=True)
    file_path = db.Column(db.String(64), index=True, unique=True)
    date = db.Column(db.DateTime, nullable=False, default=datetime.now())


# DBの生成
@app.cli.command('initialize_DB')
def initialize_DB():
    db.create_all()


# トップ画面へのルーティング
@app.route('/')
def index():
    title = 'トップ画面'
    if not session.get('login'):
        return redirect(url_for('login'))
    else:
        reg_data = DB.query.all()
        return render_template('index.html', title=title, reg_data=reg_data)


# ログイン画面へのルーティング
@app.route('/login')
def login():
    title = 'ログイン画面'
    return render_template('login.html', title=title)


# 入力情報の確認
@app.route('/logincheck', methods=['POST'])
def logincheck():
    id = request.form['id']
    password = request.form['password']
        
    if id in ID_PW:
        if password == ID_PW[id]:
            session['login'] = True
        else:
            session['login'] = False
    else:
        session['login'] = False

    if session['login']:
        return redirect(url_for('index'))
    else:
        flash('もう一度 ID とPW を正しく入力してください')
        return redirect(url_for('login'))


# ログアウト
@app.route('/logout')
def logout():
    session.pop('login', None)
    return redirect(url_for('index'))


# アップロード画面へのルーティング
@app.route('/upload')
def upload():
    title = 'アップロード画面'
    return render_template('upload.html', title=title)


@app.route('/upload_register', methods=['POST'])
def upload_register():
    title = request.form['title']
    if title:
        file = request.files['file']
        file_path = 'static/' + secure_filename(file.filename)
        file.save(file_path)
        # Excelファイルを編集してデータベースに保存する
        file_path = excelpy(file_path)
        register_data = DB(title=title, file_path=file_path)
        db.session.add(register_data)
        db.session.commit()
        flash('ファイルが用意できました')
        return redirect(url_for('index'))
    else:
        flash('タイトルを入力してもう一度アップロードしてください')
        return redirect(url_for('index'))


# 編集されたExcelファイルのダウンロード
@app.route('/download/<int:id>')
def download(id):
    download_data = DB.query.get(id)
    download_file = download_data.file_path[7:]
    print(download_file)
    return send_from_directory('static', download_file, as_attachment=True)


# ExcelファイルとDB情報の削除
@app.route('/delete/<int:id>')
def delete(id):
    delete_data = DB.query.get(id)
    delete_file = delete_data.file_path
    db.session.delete(delete_data)
    db.session.commit()
    os.remove(delete_file)
    flash('ファイルを削除しました')
    return redirect(url_for('index'))


# アプリケーションの起動
if __name__ == '__main__':
    app.run(debug=True)
