import os
import time
import threading
import webbrowser
import msal
from flask import Flask, request
from werkzeug.serving import make_server

class TokenAcquisition:
    def __init__(self, tenant_id, client_id, client_secret, scopes):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scopes = scopes
        self.authority = f'https://login.microsoftonline.com/{tenant_id}'
        self.redirect_path = "/getAToken"
        self.redirect_uri = f"http://localhost:8000{self.redirect_path}"
        self.flask_app = Flask(__name__)
        self.server = None
        self.access_token = None
        self.msal_app = msal.ConfidentialClientApplication(client_id, authority=self.authority, client_credential=client_secret)
        
        self.flask_app.add_url_rule(self.redirect_path, view_func=self.authorized, methods=['GET'])
    
    def run_server(self):
        self.server = make_server('localhost', 8000, self.flask_app)
        self.server.serve_forever()
        
    def authorized(self):
        code = request.args.get('code')
        if code:
            result = self.msal_app.acquire_token_by_authorization_code(code, scopes=self.scopes, redirect_uri=self.redirect_uri)
            if "access_token" in result:
                self.access_token = result['access_token']
                response = """
                <html>
                    <body>
                        <script>
                            window.setTimeout(function(){
                                window.close();
                            }, 3000);
                        </script>
                    </body>
                </html>
                """
                threading.Thread(target=lambda: self.server.shutdown()).start()
                return response
            else:
                return "Failed to acquire token."
        else:
            return "No code found in request."
    
    def acquire_token(self):
        auth_url = self.msal_app.get_authorization_request_url(self.scopes, redirect_uri=self.redirect_uri)
        print(f"Please authorize: {auth_url}")
        webbrowser.open(auth_url)
        server_thread = threading.Thread(target=self.run_server)
        server_thread.start()
        while self.access_token is None:
            time.sleep(2)
        return self.access_token

# Usage
if __name__ == "__main__":
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    scopes = ["Chat.ReadWrite"]
    
    token_acquisition = TokenAcquisition(tenant_id, client_id, client_secret, scopes)
    access_token = token_acquisition.acquire_token()
    print("Access token found:", access_token)
    # You can now use the access_token for further operations
