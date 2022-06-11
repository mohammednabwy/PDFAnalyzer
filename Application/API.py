#----------------------Libraries--------------------------------
import os
from flask import Flask, request, jsonify,send_file
from werkzeug.utils import secure_filename
from flasgger.utils import swag_from
from flasgger import Swagger,LazyJSONEncoder
import fitz
import json
import shutil
import requests
#--------------------Public Variables----------------------------
UPLOAD_FOLDER = 'UploadFiles/'
#UPLOAD_FOLDER_LOs=UPLOAD_FOLDER + 'LOs/'
ALLOWED_EXTENSIONS = {'doc', 'docx', 'pdf'}
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
APP = Flask(__name__)
#----------------------------------------------------------------
def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
#------------------------Swagger----------------------------------
def initSwaggerUI(app):
    app.config["SWAGGER"] = {"title": "Swagger-UI", "uiversion": 3 ,"description": "DocumnetsMLOsExtraction v1.1 API"}
    swagger_config = {
    "headers": [],
    "specs": [
        {
            "endpoint": "DocumnetsMLOsExtraction",
            "route": "/swagger.json",
            "rule_filter": lambda rule: True,  # all in
            "model_filter": lambda tag: True,  # all in
        }
    ],
    "static_url_path": "/flasgger_static",   
    "swagger_ui": True,
    "specs_route": "/swagger/",
    }
#    template = dict(swaggerUiPrefix=LazyString(lambda: request.environ.get("HTTP_X_SCRIPT_NAME", "")),securityDefinitions= {"APIKeyHeader": {"type": "apiKey", "name": "api-key", "in": "header"}}, schemes = [
#    "http",
#    "https"
#  ])
    app.json_encoder = LazyJSONEncoder
    #swagger = Swagger(app, config=swagger_config, template=template)  
    swagger = Swagger(app, config=swagger_config) 
    return swagger
#----------------------EndPoints----------------------------------
@app.route('/api/v1/ConvertToWord', methods=["POST"])
@swag_from("swagger_config_files/WordConversion.yaml")
def ConvertToWord(): 
    try:
        try:        
            file = request.files['file']
        except:
             return jsonify({'message' : 'No file'}), 406 #HTTP_406_NOT_ACCEPTABLE             
        if not allowed_file(file.filename): 
            return jsonify({'message' : 'Not allowed file extension'}), 406 #HTTP_406_NOT_ACCEPTABLE        
        filename = secure_filename(file.filename)
        filePath=os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filePath)  
        #--------------------------
        #TODO	  	 
        output_file_path=''
        if '.doc' in filename or '.docx' in filename:
             output_file_path=filePath            
        elif '.pdf' in filename:
            wordfilename=filename.split('.')[0]+".pdf"		  
            out_docx_file_path=os.path.join(app.config['UPLOAD_FOLDER'], wordfilename)
            wordsCount = get_text_words_count(filePath)
            print("wordsCount = ", wordsCount)
            if wordsCount == 0:
               print("fully scanned PDF - no relevant text")
               convertScannedPDFToWord(filePath,out_docx_file_path)
               output_file_path=out_docx_file_path              
            else:
                print("not fully scanned PDF - text is present")               
                convertEditablePDF2Word(filePath,out_docx_file_path)
                output_file_path=out_docx_file_path                             
        #--------------------------
        output_file_path=filePath
        output_file_path=output_file_path.replace('\\','/')            
        return send_file(output_file_path, as_attachment=True)  
    except Exception as e:   
         print(e)
         return jsonify({'message' : 'Internal Server Error'}), 500
#--------------------------------------------------------  
def get_text_words_count(file_name):
    doc = fitz.open(file_name)
    wordsCount=0
    for page_num, page in enumerate(doc):
        pageWordlist = page.getTextWords() 
        wordsCount=wordsCount +  len(pageWordlist)       
    doc.close()
    return wordsCount
#--------------------------------------------------------  
def convertEditablePDF2Word(pdf_file_path,out_docx_file_path):
  from pdf2docx import Converter
  # convert pdf to docx
  cv = Converter(pdf_file_path)
  cv.convert(out_docx_file_path)      # all pages by default
  cv.close() 
  return  out_docx_file_path
#-------------------------------------------------------- 
def convertScannedPDFToWord(pdf_file_path,out_docx_file_path) :
    # Provide your username and license code
    UserName = 'mnabwy'
    LicenseCode =  '865239E3-F6A4-4204-8BB1-DA7E1A448020'
    # Convert first 5 pages of multipage document into doc and txt
    RequestUrl = 'http://www.ocrwebservice.com/restservices/processDocument?language=english&pagerange=1-5&outputformat=doc,txt';

    with open(pdf_file_path, 'rb') as image_file:
        image_data = image_file.read()        
    response = requests.post(RequestUrl, data=image_data, auth=(UserName, LicenseCode))
    print(response.status_code)
    # Decode Output response
    jobj = json.loads(response.content)
    #Download output file (if outputformat was specified)
    file_response = requests.get(jobj["OutputFileUrl"], stream=True)
    with open(out_docx_file_path, 'wb') as output_file:
       shutil.copyfileobj(file_response.raw, output_file)
    return out_docx_file_path
#--------------------------------------------------------
if __name__ == '__main__':
    swagger=initSwaggerUI(app)
    app.run(host='127.0.0.1', port=5000)   
#--------------------------------------------------------  