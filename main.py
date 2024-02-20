import time
from flask import Flask, request
import msal
import webbrowser
import threading
import os
import requests
from werkzeug.serving import make_server

# Read environment variables
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')

# Azure AD app registration details
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_PATH = "/getAToken"
REDIRECT_URI = f"http://localhost:8000{REDIRECT_PATH}"

# MSAL configuration
SCOPES = ["Chat.ReadWrite"]  # Add other scopes/permissions as needed
msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)

# Flask app to handle the redirect and acquire the token
flask_app = Flask(__name__)

# Global variable to store the access token
access_token = None

server = None  # Define a variable for the server instance

# Modify the authorized route to signal the server to stop after getting the token
@flask_app.route(REDIRECT_PATH)
def authorized():
    global access_token, server
    code = request.args.get('code')
    if code:
        result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
        if "access_token" in result:
            access_token = result['access_token']

            # Prepare a response with JavaScript to display a message and close the window
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
            
            # Use a separate thread to shut down the server after sending the response
            if server is not None:
                threading.Thread(target=lambda: server.shutdown()).start()
            
            return response
        else:
            return "Failed to acquire token."
    else:
        return "No code found in request."

def run_server():
    global server
    server = make_server('localhost', 8000, flask_app)
    server.serve_forever()

def get_chats(access_token):
    """
    Retrieve the chats for the authenticated user.

    If the access token is invalid (e.g., expired), this function will re-run the process to get a new access token.

    Args:
        access_token (str): The access token to authenticate the request.

    Returns:
        list: A list containing the chats for the authenticated user.
    """
    url = "https://graph.microsoft.com/v1.0/me/chats"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    
    response.raise_for_status()
    return response.json()['value']


def get_last_message_of_chat(access_token, chat_id):
    """
    Retrieve the last 10 messages of a particular chat.

    Args:
        access_token (str): The access token to authenticate the request.
        chat_id (str): The ID of the chat from which to retrieve messages.

    Returns:
        list: A list containing the last 10 messages of the specified chat.
    """
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
    headers = {'Authorization': f'Bearer {access_token}'}
    params = {
        '$top': 1,  # Limit the response to the last 10 messages
        '$orderby': 'createdDateTime DESC'  # Order by createdDateTime in descending order
    }
    response = requests.get(url, headers=headers, params=params)
    
    response.raise_for_status()  # This will raise an exception for HTTP error responses
    return response.json()['value']

def main():
    # Generate the authorization URL and open it in the browser
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    print(f"Please authorize: {auth_url}")
    webbrowser.open(auth_url)

    # Start the Flask server in a separate thread
    server_thread = threading.Thread(target=run_server)
    server_thread.start()

    # Main thread waits until the access token is acquired
    while access_token is None:
        time.sleep(2)
    
    # Server shutdown is handled in the authorized route once the token is acquired

    # for chat in get_chats(access_token):
    #     print(f"Chat Topic: {chat['topic']} Id: {chat['id']}, Chat Type: {chat['chatType']}")
    result = get_last_message_of_chat(access_token, "19:00c83d8a7ff4451abdd883041d01f9e4@thread.v2")
    print(result[0]['body']['content'])

if __name__ == "__main__":
    main()