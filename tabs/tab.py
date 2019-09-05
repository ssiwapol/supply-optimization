import os

from flask import request, send_file

import mod

p = mod.PathFile()
configinfo = p.configinfo()


def set_callbacks(app):
    @app.server.route('/download/<file>')
    def download_output(file):
        p.setuser(request.authorization['username'])
        if file == "input_template":
            filename = "input_template.xlsx"
            pathfile = os.path.join("./models", filename)
        else:
            filename = configinfo['file'][file]
            pathfile = p.loadfile(filename)
        return send_file(pathfile,
                         attachment_filename=filename,
                         as_attachment=True, cache_timeout=0)