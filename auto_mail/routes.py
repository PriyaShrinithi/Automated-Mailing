import pandas as pd
from flask import render_template, redirect, url_for, flash, request
from auto_mail import app, mail
from flask_mail import Message

def get_recipient_list(form_file):
    sheets = pd.read_excel(form_file, sheet_name=None)
    recipient_list = []
    for sheet, data in sheets.items():
        recipient_list += data['EMAIL'].values.tolist()
    return recipient_list

@app.route('/', methods=['GET', 'POST'])
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method=='POST':
        subject = request.form.get('subject')
        body = request.form.get('body')
        file = request.files['file']
        recipient_list = get_recipient_list(file)
        print(recipient_list)
        msg = Message(subject, sender='dwight@dundermifflin.com', recipients=recipient_list)
        msg.body = body
        mail.send(msg)
        flash('Mail has beenn sent!', 'success')
        return '<center>The mails have been sent!</center>'
    return render_template('upload.html')

