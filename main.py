from openpyxl import load_workbook
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename
from shutil import copyfile
import os 
import time
'''
TODO: Excel row adj, cancel button cancel, delete all things in target dir before restarting
'''


app = Flask(__name__) 


ALLOWED_EXTENSIONS = set(['ipt', 'xlsx'])
app.config['ALLOWED_EXTENSIONS'] = ALLOWED_EXTENSIONS

app.config['UPLOAD_FOLDER'] = 'static/file_queue/'


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


@app.route('/', methods = ['GET', 'POST'])
def home():
    folder = 'static/file_queue/'
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            #elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e)
    return render_template('index.html')



@app.route('/finished', methods = ['GET', 'POST'])
def sortedPage():
    if request.method == 'POST':
        uploaded_files = request.files.getlist("file[]")
        filenames = []
        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                filenames.append(filename)
        return redirect(url_for('doTasks'))


@app.route('/task')
def doTasks():
    xlFileCount = 0
    iptFileCount = 0

    files = os.listdir('static/file_queue/')
    for i in files:
        if ".xlsx" in i:
            xlFile = i
            xlFileCount+=1
        elif ".ipt" in i:
           iptFileCount+=1
        else:
            pass
    if(iptFileCount == 0):
        return render_template('error.html', err="Status: No .ipt files have been uploaded! Please restart.")
    
    elif(xlFileCount > 1 or xlFileCount == 0):
        return render_template('error.html', err="Multiple or none excel files uploaded! Please restart.")
    else:
        readxl(xlFile)
        return render_template('success.html')


def readxl(excelFileName):
    wb = load_workbook(filename='static/file_queue/' + excelFileName)#, read_only=True)
    #Change according to sheet name
    ws = wb['Sheet1']
    for row in ws.rows:
        myrow = [] 
        for cell in row: 
            myrow.append(cell.value)
        if (myrow[5] == "Yes"):
            part_num = myrow[2]
            material = myrow[3]
            #print("Part: " + part_num + " Material: " + material)
            dst = "material/" + material + "/" + part_num + ".ipt"
            if not os.path.exists("material/"): 
                os.mkdir("material/")
            src = "static/file_queue/" + part_num + ".ipt"
            if not os.path.exists("material/"+ material): 
                os.mkdir("material/"+material)
            copyfile(src, dst)
    #Find a way to close the file and be able to delete the excel file from the saved location (file_queue)
    #wb.close_workbook()
    
    






if __name__ == '__main__':
    app.run(host='0.0.0.0', port=69, debug=True)
    #readxl("blueprint.xlsx")