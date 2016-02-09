import os
from flask.ext.bootstrap import Bootstrap
from flask.ext.wtf import Form
from wtforms import FileField, SubmitField, ValidationError
from flask import (Flask, render_template, request, 
                    make_response,redirect, url_for, Markup)
import functions as fns
import json
app = Flask(__name__)
app.config['SECRET_KEY'] = 'top secret!'
bootstrap = Bootstrap(app)

class UploadForm(Form):
    image_file = FileField('Archivo de datos')
    submit = SubmitField('Cargar')

    def validate_image_file(self, field):
        if ((field.data.filename[-5:].lower() != '.xlsm') 
        	and (field.data.filename[-5:].lower() != '.xlsx')):
            raise ValidationError('Invalid file extension')





@app.route('/', methods=['GET', 'POST'])
def index():
    image = None
    form = UploadForm()
    fullfilename = None
    if form.validate_on_submit():
        image = 'uploads/' + form.image_file.data.filename
        fullfilename = os.path.join(app.static_folder, image)
        form.image_file.data.save(fullfilename)
        response = make_response(render_template('index.html', form=form, image=image,fullfilename=fullfilename))
        response.set_cookie('filename',json.dumps({'file': fullfilename}))
        return response
    else:
        return render_template('index.html', form=form, image=image,fullfilename=fullfilename)

@app.route('/save',methods=['POST','GET'])
def save():
    response = make_response(redirect(url_for('index')))
    response.set_cookie('filename',json.dumps(dict(request.form.items())))
    return response

@app.route('/process', methods=['GET', 'POST'])
def process():
    filecookie = request.cookies.get('filename')
    # import pdb as pdb
    # pdb.set_trace()
    fullfilename = json.loads(filecookie)['file'] 
    grantitulo, memotext, pdftext, table_list, plot_list, title_list, description_list, nlen = fns.reporter(fullfilename) 

    return render_template('Output.html', grantitulo=grantitulo, memotext=memotext,pdftext=pdftext,
        table_list=table_list,plot_list=plot_list,
        title_list=title_list,description_list=description_list, nlen = nlen)



if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
app.run(host='0.0.0.0', port=port,debug=True)

	