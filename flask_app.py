import os

from flask import Flask, render_template, url_for, request, redirect
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime


# basedir = 'G:\\Projects\\site_src_python\\blog.sqlite'
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
    intro = db.Column(db.String(300), nullable=False)
    text = db.Column(db.Text, nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return '<Article %r>' % self.id

class Visitor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(100), nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)




@app.route('/user/<name>')
def user(name):
    return '<h1>Hello, %s !</h1>'%name

@app.route('/')
def hello_world():
    return 'Hello World on http://zelenskiy.pythonanywhere.com/!'

@app.route('/create',methods=['POST', 'GET'])
def create_article():
    if request.method == "POST":
        print(request.user_agent)
        print(request.remote_addr)
        title = request.form['title']
        intro = request.form['intro']
        text = request.form['text']
        article = Article(title=title, intro=intro, text=text)
        try:
            db.session.add(article)
            db.session.commit()

            return redirect('/')
        except:
            return "error"

    else:
        return render_template("create-article.html")





if __name__ == '__main__':
    app.run(debug=True)
