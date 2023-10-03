from flask import Flask,request
import win32com.client
from flask_cors import CORS
import os
import random
import pythoncom
import codecs
from bs4 import BeautifulSoup
import glob

from cryptography.fernet import Fernet
import pdfkit

app= Flask(__name__)
CORS(app)

""" API for Conversion of Doc to HTML """

@app.route('/doctohtml', methods=['GET','POST'])
def doctohtml():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            # print("File path",filepath) #Only to be uncommented in case of testing
            pythoncom.CoInitialize()
            doc = win32com.client.GetObject((filepath))
            num=random.randint(100000,999999)
            doc.SaveAs (FileName=filepath+str(num)+".html", FileFormat=8)
            doc.Close()
            return f' Html file generated and saved successfully with name { filepath }{ str(num) }.html'
            # f=codecs.open(filepath+str(num)+'.html', 'r')
            # document= BeautifulSoup(f.read()).get_text()
            # return document
            # with open(filepath+str(num)+'.html','r') as path:
            #     return path
        
        else:
            return f'Kindly trigger API using POST method'
    except Exception as e:
        print(e)



port = int(os.getenv('PORT', 8080)) 

# run 
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port, debug=False) # deploy with debug=False