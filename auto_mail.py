import pandas as pd
from flask import Flask
from flask_mail import Message, Mail
import os

app = Flask(__name__)

app.config['MAIL_SERVER'] = 'smtp.googlemail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
#create an environment variable where EMAIL_USER is email-id and EMAIL_PASS is the password
app.config['MAIL_USERNAME'] = os.environ.get('EMAIL_USER')
app.config['MAIL_PASSWORD'] = os.environ.get('EMAIL_PASS')


def get_recipient_list():
    sheets = pd.read_excel("form.xlsx", sheet_name=None)
    recipient_list = []
    for sheet, data in sheets.items():
        recipient_list+=data['EMAIL'].values.tolist()
    return recipient_list

@app.route('/')
def send_mail():
    recipient_list = get_recipient_list()
    mail = Mail(app)
    msg = Message('Hello There', sender='dwight@dundermifflin.com', recipients=recipient_list)
    msg.body = f'''Hello. This is Dwight.
I shall count from 1 to 10. Please hide your persons and escape the murderer.
Else you will die by the hands of a Schrute. If you do not want to hide, meet me at
https://meet.google.com/hkq-xxrw-fcr and you shall be safe.
'''
    mail.send(msg)
    return '<center>sent!</center>'

if __name__ == "__main__":
    app.run(debug=True)
