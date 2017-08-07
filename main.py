from openpyxl import load_workbook
from flask import Flask, request, redirect, url_for, render_template
from shutil import copyfile
import os 
import webbrowser

url = 'http://localhost:69/'

app = Flask(__name__)
webbrowser.open(url) 

ALLOWED_EXTENSIONS = set(['ipt', 'xlsx'])
app.config['ALLOWED_EXTENSIONS'] = ALLOWED_EXTENSIONS

file_queue = 'D:/Ibex Innovation/File Sorter/file_queue/'
materials =  'D:/Ibex Innovation/File Sorter/materials/'

if not os.path.exists(file_queue): 
                os.makedirs(file_queue)

app.config['UPLOAD_FOLDER'] = file_queue

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods = ['GET', 'POST'])
def home():
    for the_file in os.listdir(file_queue): 
        file_path = os.path.join(file_queue, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(e)
    if not os.path.exists(materials): 
                os.mkdir(materials)
    for the_file in os.listdir(materials):
         file_path = os.path.join(materials, the_file)
         try:
             if os.path.isfile(file_path):
                 os.unlink(file_path)
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
                filename = file.filename
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                filenames.append(filename)
        return redirect(url_for('doTasks'))

@app.route('/task')
def doTasks():
    xlFileCount = 0
    iptFileCount = 0

    files = os.listdir(file_queue)
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
        return render_template('error.html', err="Status: Multiple or none .xlsx files were uploaded! Please restart.")

    if(iptFileCount == 0 & xlFileCount == 0):
        return render_template('error.html', err="Status: No .ipt or .xlsx files were uploaded! Please restart.")

    else:
        readxl(xlFile)
        return render_template('success.html')
        exit()

def readxl(excelFileName):
    wb = load_workbook(filename=file_queue + excelFileName)
    ws = wb['Sheet1']
    for row in ws.rows:
        myrow = [] 
        for cell in row: 
            myrow.append(cell.value)
        if (myrow[5] == "Yes"):
            part_num = myrow[2]
            material = myrow[3]
            dst = materials + material + "/" + part_num + ".ipt"
            if not os.path.exists(materials): 
                os.mkdir(materials)
            src = file_queue + part_num + ".ipt"
            if not os.path.exists(materials + material): 
                os.mkdir(materials+material)
            copyfile(src, dst)

@app.route('/close', methods = ['GET', 'POST'])
def exit():
    os._exit(0)

    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=69, debug=False)

    
    
