import os
os.environ['FLASK_ENV_LABEL'] = 'DEV'
from web_app.app import create_app
app = create_app()
app.run(host='127.0.0.1', port=5099, debug=False)
