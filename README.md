The apis.py contains 2 API's 

doctohtml : The doctohtml API converts the Documents of both extension types doc & docx into html format.
            The API takes filepath as an input with requesting type method as POST and creates HTML file in the respective filepath location.

pdftodoc :  The pdftodoc API converts the PDF's into docx format.
            The API takes filepath as an input with requesting type method as POST and creates Docx file in the respective filepath location.

doctopdf :  The doctopdf API converts the Docs into pdf format.
            The API takes filepath as an input with requesting type method as POST and creates PDF file in the respective filepath location.

htmltopdf : The htmltopdf API converts the HTML into pdf format.
            The API takes filepath as an input with requesting type method as POST and creates PDF file in the respective filepath location.

keygen : This API is used to generate filekey which can be used for both encryption and
         Decryption.

encryptor : This API takes an input filepath and encrypt the file using the key generated 
            via keygen API.

decryptor : This API takes an input filepath and encrypt the file using the key generated 
            via keygen API.


Note: Install Ms-Office on the machine for document conversion.

      Download and Install wkhtml from the link : https://wkhtmltopdf.org/downloads.html
      After installing setup the environment path variable for wkhtml e.g for above case it will be C:\Program Files\wkhtmltopdf\bin
