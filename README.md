# CNC-File-Sorter-Flask-HTML
A file sorter that is capable of checking for CNC files from an excel spreadsheet generated by Autodesk Inventor. Then the program sorts CNC file materials and creates folders based on materials of .ipt files and pastes the files to those individual folders.

# This is not a production level software. 
# This guide is for windows machines(feel free to implement it to other operating systems).

<h2> Programming languages used: </h2>
<ul>
  <li>HTML</li>
  <li>CSS</li>
  <li>JavaScript</li>
  <li>Pyhton</li>
</ul>

<h2>Technologies used:</h2>
<ul>
  <li>Bootstrap</li>
  <li>Flask</li>
  <li>JQuery</li>
  <li>Ajax</li>
  <li>Openpyxl</li>
</ul>

<h2>To setup:</h2>

<p><strong>#1.</strong> Download and install <code><a href="https://www.python.org/downloads/release/python-340/" target="about_blank">Python 3.4</a></code> (Make sure this is the version. I found this version to be the most stable to run the script. Feel free to play around with other versions and see if the program works.)</p>

<p><strong>#2.</strong> Edit the environment variables: <br>
&nbsp;&nbsp;&nbsp;A. Using the windows search, type <code>"edit the system environment variables"(not edit the system environment variables for your account)</code> and hit enter.<br>
&nbsp;&nbsp;&nbsp;B. When the window opens up, click <code>Environment Variables</code><br>
&nbsp;&nbsp;&nbsp;C. When the next window opens up, under <code>"System Variables"</code>, look under the <code>Variable</code> column for <code>Path</code> and select it. &nbsp;&nbsp;Under the <code>"System Variables"</code> box, there are buttons, look for and select <code>"Edit"</code>.</br>
&nbsp;&nbsp;&nbsp;D. In the next window that pops up, look for the <code>New</code> button. In the blank, add the location of the pip script. To find the &nbsp;&nbsp;&nbsp;location of the pip script:<br> &nbsp;&nbsp;&nbsp;&nbsp;a. Go to the location where you installed python. <br>&nbsp;&nbsp;&nbsp;&nbsp;b. Inside the folder look for folder named <code>Script</code> and hit enter. <br>&nbsp;&nbsp;&nbsp;&nbsp;c. Copy the location of this <code>Script</code> folder<br>&nbsp;&nbsp;&nbsp;Then paste the location in the blank. Hit <code>Ok</code> on all of the windows that poped up.</p>

<p><strong>#3.</strong> Intstall the dependencies using pip<br>
&nbsp;&nbsp;&nbsp;A. Go to windows search and type <code>"cmd"</code> and when <code>Command Prompt</code> shows up in the search menu, right click and hit <code>Run &nbsp;&nbsp;&nbsp;as administrator</code>. in the next propmpt, click <code>Yes</code><br>
&nbsp;&nbsp;&nbsp;B. When Command Prompt pops up, it should say <code>C:\WINDOWS\system32></code>.<br>
&nbsp;&nbsp;&nbsp;C. Next we need to install a module to run the Python script called <code><a href="http://flask.pocoo.org/" target="about_blank">flask</a></code>. This module will handle the <code>GET</code>, <code>POST</code> &nbsp;&nbsp;&nbsp;requests and redirects. To install it:<br>&nbsp;&nbsp;&nbsp;&nbsp;a. Type <code>"pip install pyinstaller"</code> and hit enter in the Command Prompt window. This package will take care of &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;installing the <code>flask</code> module and others.<br>
&nbsp;&nbsp;&nbsp;&nbsp;b. Once it is done installing, type <code>"pip install flask"</code><br>&nbsp;&nbsp;&nbsp;Now flask will be installed on your computer.<br>
&nbsp;&nbsp;&nbsp;D. Now you need to install the last module called <code><a href="https://openpyxl.readthedocs.io/en/default/" target="about_blank">Openpyxl</a></code> which will read the excel spreadsheets generated by &nbsp;&nbsp;&nbsp;AutoDesk Inventor. To install it:<br>
&nbsp;&nbsp;&nbsp;&nbsp;a. In the same Command Prompt window, type <code>"pip install openpyxl"</code> and hit enter<br>
&nbsp;&nbsp;&nbsp;Now you have all of the dependencies needed to run the Pyhton script.</p>

<p><strong>#4.</strong> Download this github repository by clicking the green download button.</p>

<p><strong>#5.</strong> Extract the downloaded .zip file into a drive other than the main <code>C:</code> drive. This is because windows has a high level &nbsp;&nbsp;&nbsp;&nbsp;security in place which does not permit the program to create folders and store files in them.</p>

<p><strong>#6.</strong> Open the Pyhton <code>main.py</code> script using a text editor or IDLE. Then:<br>
&nbsp;&nbsp;&nbsp;&nbsp;a. Edit the variable named <code>file_queue</code> value so that the drive letter of the location of the value matches the drive letter &nbsp;&nbsp;&nbsp;&nbsp;where you extracted the .zip file on. <br>&nbsp;&nbsp;&nbsp;&nbsp;For example:<br>
&nbsp;&nbsp;&nbsp;&nbsp;This is the original code:<br>
&nbsp;&nbsp;&nbsp;&nbsp;<code>file_queue = 'D:/File Sorter/file_queue/'</code><br>
&nbsp;&nbsp;&nbsp;&nbsp;If you extracted the .zip file on <code>F:</code> drive, your changes should look like this:<br>
&nbsp;&nbsp;&nbsp;&nbsp;<code>file_queue = 'F:/File Sorter/file_queue/'</code><br>
&nbsp;&nbsp;&nbsp;&nbsp;Do the same for materials variable.</p>

<p><strong>#7.</strong> Now you can locate the <code>main.py</code> script and hit enter to run the program.</p> 


