# PANDORE SNIFFER API - Flask application

# IMPORTS======================================================================

from flask import Flask
from flask_swagger_ui import get_swaggerui_blueprint

# FLASK APP=====================================================================

sniffer_gui = Flask(__name__)

# SWAGGER=======================================================================
SWAGGER_URL = '/api'
API_URL = '/static/swagger/swagger.json'
SWAGGERUI_BLUEPRINT = get_swaggerui_blueprint(
    SWAGGER_URL,
    API_URL,
    config={
        'app_name': "Pandore Sniffer API"
    }
)
sniffer_gui.register_blueprint(SWAGGERUI_BLUEPRINT, url_prefix=SWAGGER_URL)

# VIEWS=========================================================================

from gui.views import *