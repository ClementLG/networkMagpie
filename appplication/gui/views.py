# PANDORE SNIFFER API - Flask views

# IMPORTS======================================================================

from flask import Flask, jsonify, request, render_template, Response
from gui import sniffer_gui
from gui.functions import *


# FLASK ROUTES===================================================================

@sniffer_gui.route("/", methods=["GET"])
@sniffer_gui.route('/index', methods=["GET"])
def index():
    return render_template("index.html")