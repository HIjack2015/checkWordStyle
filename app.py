import os

from flask import Flask, request, send_from_directory, render_template

# set the project root directory as the static folder, you can set others.
from main import check_all

app = Flask(__name__, static_url_path='')

@app.route('/static/<type>/<path:path>')
def send_js(type,path):
    return app.send_static_file(os.path.join(type, path).replace('\\', '/'))

@app.route('/index')
def index():
    return render_template("index.html")



@app.route('/uploader', methods=['GET', 'POST'])
def upload_file():
    try:
        if request.method == 'POST':

            f = request.files['file']
            file_path ="data/"+  (f.filename)
            f.save(file_path)

            res=check_all(file_path)
            return render_template("result.html",res=res)
    except Exception as e:
        return render_template("result.html", res=[[str(e)]])
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80)