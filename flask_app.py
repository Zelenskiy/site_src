#! /usr/bin/env python
# -*- coding: utf-8 -*-
"""
Процес створення google Ф

"""

import os
from sheetutils import addBlock, searchOnDate, test

from flask import Flask, render_template, request, redirect
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

# basedir = 'G:\\Projects\\site_src_python\\blog.sqlite'
from tt import pzCreate

basedir = os.path.abspath(__file__)



app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///blog.sqlite'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///'+'G:\\Projects\\site_src_python\\blog.sqlite'
# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///'+os.path.join(basedir, 'blog.sqlite')
db = SQLAlchemy(app)


class Article(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(100), nullable=False)
    geo = db.Column(db.String(300), nullable=False)
    text = db.Column(db.Text, nullable=False)
    # ip_address = db.Column(db.String(40), default='0.0.0.0', nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)

    # ip = db.Column(db.ip)

    def __repr__(self):
        return '<Article %r>' % self.id


class Visitor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ip_address = db.Column(db.String(40), default='0.0.0.0')
    title = db.Column(db.String(100), nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)


@app.route('/user/<name>')
def user(name):
    return '<h1>Hello, %s !</h1>' % name

@app.route('/readgss')
def readGoogleSpreadSheets():
    s = test()
    return s


@app.route('/')
def hello_world():
    return 'Hello World on http://zelenskiy.pythonanywhere.com/!'\

@app.route('/createpz/', methods=['POST'])
def createpz():
    if request.method == "POST":
        d1 = request.form['date_start']
        d2 = request.form['date_finish']
        idMissTable = request.form['idmisstable']
        nameSheetMissTable = request.form['namesheetmisstable']
        # urlPZ = nameSheetMissTable
        urlPZ = pzCreate(idMissTable, nameSheet=nameSheetMissTable, d1=d1, d2=d2)
        # Додаємо адресу до таблиці


        return 'ПОЯСНЮВАЛЬНИ ЗАПИСКА<BR>' \
               'за період від ' + d1 + ' до ' + d2 + '. <br>' \
               'Згенеровано ' + str(datetime.now()) + '. <br>' \
               'Покликання на сторінку <a href="' + urlPZ + '">' + urlPZ + '</a>'


    return 'Hello World on http://zelenskiy.pythonanywhere.com/!'

@app.route('/addblock/', methods=['POST'])
def add_block():
    url_my = ''
    if request.method == "POST":
        text = request.form['text']
        url_my = request.form['url_client']
        idSpreadheet = request.form['idmisstable0']
        p = (text.split("\r\n"))
        lst = []
        for r in p:
            w = r.split("\t")
            tmp = []
            for c in w:
                tmp.append(c)
            lst.append(tmp)
        print(url_my)
        addBlock(idSpreadheet, "missingbook", lst)

    return redirect("https://znz16300.github.io/site/zamini.html")


@app.route('/posts')
def posts():
    articles = Article.query.order_by(Article.date).all()
    return render_template('posts.html', articles=articles)


@app.route('/create', methods=['POST', 'GET'])
def create_article():
    if request.method == "POST":
        print(request.user_agent)
        title = request.form['title']
        geo = request.form['geo']
        text = request.form['text']

        article = Article(title=title, geo=geo, text=text)
        try:
            db.session.add(article)
            db.session.commit()

            return redirect('/create')
        except:
            return "error"

    else:
        return render_template("create-article.html")


if __name__ == '__main__':
    app.run(debug=True)
