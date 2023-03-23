import pandas as pd
from flask import Flask, render_template, request, send_file
import os

app = Flask(__name__)


@app.route('/')
def upload():
    # list all files in model folder
    files = os.listdir('model')
    return render_template('upload-excel.html', files=files)


@app.route('/download/<filename>')
def download(filename):
    # send file from model folder as attachment
    return send_file(os.path.join('model', filename), as_attachment=True)


@app.route('/read-excel', methods=['POST'])
def read_excel():
    file = request.files['file']
    file2 = request.files['file2']
    ratio_normal = request.form["ratio_normal"]
    ratio_work = request.form.get("ratio_work")
    ratio_check = request.form.get("ratio_check")
    ratio_exam = request.form.get("ratio_exam")
    # create a data folder if it does not exist
    if not os.path.exists('data'):
        os.makedirs('data')

    # save the file to data folder with its original name
    filename = file.filename
    filename2 = file2.filename
    file.save(os.path.join('data', filename))
    file2.save(os.path.join('data', filename2))

    # read the file from data folder
    df = pd.read_excel(os.path.join('data', filename), index_col=0)

    # multiply '作业平均分' and '考勤平均分' by ratio variables
    df['作业平均分'] *= ratio_work/100
    df['考勤平均分'] *= ratio_check/100
    df['考试评分']
    # calculate the sum and put it in '平时分' column
    df['平时分'] = df['作业平均分'] + df['考勤平均分']

    # generate a new xlsx file with '平时成绩' suffix
    new_filename = filename[:-5] + '平时分.xlsx'

    # save the new xlsx file to data folder
    df.to_excel(os.path.join('data', new_filename), index=False)
    return render_template('show-data.html', data=df.to_html())


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True, use_reloader=False, threaded=True)