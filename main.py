import time
from flask import Flask, request
import msal
import webbrowser
import threading
import os
import requests
from werkzeug.serving import make_server

from token_aquisition import TokenAcquisition

# Read environment variables
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
SCOPES = ["Chat.ReadWrite"]  # Add other scopes/permissions as needed


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
    token_acquisition = TokenAcquisition(TENANT_ID, CLIENT_ID, CLIENT_SECRET, SCOPES)
    access_token = token_acquisition.acquire_token()
    print("Access token found:", access_token)
    
    # Server shutdown is handled in the authorized route once the token is acquired

    # for chat in get_chats(access_token):
    #     print(f"Chat Topic: {chat['topic']} Id: {chat['id']}, Chat Type: {chat['chatType']}")
    result = get_last_message_of_chat(access_token, "19:00c83d8a7ff4451abdd883041d01f9e4@thread.v2")
    print(result[0]['body']['content'])

if __name__ == "__main__":
    main()