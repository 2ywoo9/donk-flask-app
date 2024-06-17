from flask import Flask, render_template, request, redirect, url_for, flash
import openpyxl

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 用于会话加密

# 加载Excel文件
wb = openpyxl.load_workbook('data.xlsx')
ws = wb.active

# 其他函数（增删查改）...

@app.route('/')
def index():
    return render_template('index.html')

# 定义路由和视图函数来处理不同的请求...

if __name__ == '__main__':
    app.run(debug=True)
