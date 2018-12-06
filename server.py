#!/usr/bin/env python
# -*- coding: utf-8 -*-


from sys import argv
import os
import bottle
from bottle import run, template, route, static_file, request, response
try:
    import win32com.client
except:
    pass


ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
STATIC = '{}/media'.format(ROOT_PATH)


def enable_cors(fn):
    def _enable_cors(*args, **kwargs):
        # set CORS headers
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Origin, Accept, Content-Type, X-Requested-With, X-CSRF-Token'

        if bottle.request.method != 'OPTIONS':
            # actual request; reply with the actual response
            return fn(*args, **kwargs)
    return _enable_cors


@route('/media/<filename>')
def server_static(filename):
    return static_file(filename, root=STATIC)


@route('/')
def home():
    return template("home.html")



@route('/send-text/', method='POST')
@enable_cors
def sendkeys():
    scan_data = request.forms.get("data", "")
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys(scan_data);
        return ">>" + scan_data
    except:
        print('>>', scan_data)
    return scan_data


_port = 8181
run(host='0.0.0.0', port=_port)
