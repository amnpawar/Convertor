from flask import Flask,request
import win32com.client
from flask_cors import CORS
import os
import random
import pythoncom
import codecs
from bs4 import BeautifulSoup
import glob

from pdf2docx import Converter
from docx2pdf import convert
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

""" API for Conversion of PDF to Doc """

@app.route('/pdftodoc', methods=['GET','POST'])
def pdftodoc():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            print("File path",filepath) #Only to be uncommented in case of testing
            cv = Converter(filepath)
            print("HERE")
            a=cv.convert(filepath+'.docx', start=0, end=None)
            cv.close()
            return f' Doc file generated and saved successfully with name { filepath }.docx'

        else:
            return f'Kindly trigger API using POST method'
    except Exception as e:
        print(e)

""" API for Conversion of Doc to PDF """

@app.route('/doctopdf', methods=['GET','POST'])
def doctopdf():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            # print("File path",filepath) #Only to be uncommented in case of testing
            pythoncom.CoInitialize()
            convert(filepath,filepath+'.pdf')
            return f' Pdf file generated and saved successfully with name { filepath }.pdf'

        else:
            return f'Kindly trigger API using POST method'
    except Exception as e:
        print(e)

""" API for Conversion of HTML to PDF """

@app.route('/htmltopdf', methods=['GET','POST'])
def htmltopdf():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            print("File path",filepath) #Only to be uncommented in case of testing
            # pythoncom.CoInitialize()
            config = pdfkit.configuration(wkhtmltopdf = r"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")
            pdfkit.from_file(filepath, filepath+'.pdf', configuration = config)
            return f' Pdf file generated and saved successfully with name { filepath }.pdf'

        else:
           return f'Kindly trigger API using POST method'
    except Exception as e:
        print(e)

@app.route('/keygen', methods=['GET','POST'])
def key_generation():
    try:
        if request.method == 'POST':
                # key generation
            key = Fernet.generate_key()

            # string the key in a file
            with open('filekey.key', 'wb') as filekey:
                filekey.write(key)
            return f'Key has been generated'

        else:
            return f'Kindly trigger API using POST method'
    except Exception as e:
        print(e)

@app.route('/encryptor', methods=['GET','POST'])
def file_encryption():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            key = Fernet.generate_key()

            # string the key in a file
            with open('filekey.key', 'wb') as filekey:
                filekey.write(key)


                # opening the key
            with open('filekey.key', 'rb') as filekey:
                key = filekey.read()

            # using the generated key
            fernet = Fernet(key)

            # opening the original file to encrypt
            with open(filepath, 'rb') as file1:
                original = file1.read()
                
            # encrypting the file
            encrypted = fernet.encrypt(original)

            # opening the file in write mode and
            # writing the encrypted data
            with open(filepath, 'wb') as encrypted_file:
                encrypted_file.write(encrypted)
            return f'File has been encrypted successfully'
        else:
            return f'Kindly trigger API using POST method'

    except Exception as e:
        print(e)

@app.route('/decryptor', methods=['GET','POST'])
def file_decryption():
    try:
        if request.method == 'POST':
            filepath = request.form['fullPath']
            with open('filekey.key', 'rb') as filekey:
                key = filekey.read()

            # using the key
            fernet = Fernet(key)
            
            # opening the encrypted file
            with open(filepath, 'rb') as enc_file:
                encrypted = enc_file.read()
            
            # decrypting the file
            decrypted = fernet.decrypt(encrypted)
            
            # opening the file in write mode and
            # writing the decrypted data
            with open(filepath, 'wb') as dec_file:
                dec_file.write(decrypted)
            return f'File has been decrypted successfully'
        
        else:
            return f'Kindly trigger API using POST method'

    except Exception as e:
        print(e)

port = int(os.getenv('PORT', 8080)) 

# run 
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port, debug=False) # deploy with debug=False