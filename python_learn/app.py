from flask import Flask, render_template, request
from script import multiply

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    error = None
    if request.method == 'POST':
        try:
            num1 = float(request.form['num1'])
            num2 = float(request.form['num2'])
            result = multiply(num1, num2)
        except Exception as e:
            error = '请输入有效的数字！'
    return render_template('index.html', result=result, error=error)

if __name__ == '__main__':
    app.run(debug=True) 