import json
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, flash, session, redirect, url_for
import os
from work_main import generate_marksheet, concise_marksheet, Send_email
import pandas as pd
from wtforms import Form, FloatField, validators

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "sample_input"
app.secret_key = "hello"

@app.route('/', methods=['GET', 'POST'])
def index():

	if request.method == "POST":
		req = request.form

		pos = req.get("correct")
		neg = req.get("wrong")

		print(pos, neg)

		if not os.path.isdir(app.config['UPLOAD_FOLDER']):
			os.mkdir(app.config['UPLOAD_FOLDER'])

		f1 = request.files["upload-file1"]
		if not os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], f1.filename)):
			f1.save(os.path.join(app.config['UPLOAD_FOLDER'], f1.filename))

		f2 = request.files["upload-file2"]
		if not os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], f2.filename)):
			f2.save(os.path.join(app.config['UPLOAD_FOLDER'], f2.filename))

		path1 = os.path.join(app.config['UPLOAD_FOLDER'], f1.filename)
		path2 = os.path.join(app.config['UPLOAD_FOLDER'], f2.filename)
		
		if request.form["action"] == "Generate Roll number wise Marksheet":
			result = generate_marksheet(path1, path2, pos, neg)

		elif request.form["action"] == "Generate Concise Marksheet with Roll Num, Obtained Marks, marks after negative":
			result = concise_marksheet(path1, path2, pos, neg)

		elif request.form["action"] == "Send Email":
			result = Send_email(path1, path2, pos, neg)
		else:
			result = None

		print(result)

		if result:
			flash("no roll number with ANSWER is present, Cannot Process!")
			return redirect(request.url)

		return redirect(request.url)
				
	return render_template('view.html')


if __name__ == '__main__':
    app.run(debug=True)