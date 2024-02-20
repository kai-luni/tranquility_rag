import json
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

def post_message_to_chat(access_token, chat_id, message_content):
    """
    Post a message to a specific chat in Microsoft Teams.

    Args:
        access_token (str): The access token to authenticate the request.
        chat_id (str): The ID of the chat where the message will be posted.
        message_content (str): The content of the message to post.

    Returns:
        dict: A dictionary representing the posted message if successful.
    """
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    body = {
        "body": {
            "content": message_content
        }
    }
    response = requests.post(url, headers=headers, data=json.dumps(body))
    
    response.raise_for_status()  # This will raise an exception for HTTP error responses
    return response.json()

def main():
    token_acquisition = TokenAcquisition(TENANT_ID, CLIENT_ID, CLIENT_SECRET, SCOPES)
    access_token = token_acquisition.acquire_token()
    print("Access token found:", access_token)

    chat_id = "19:00c83d8a7ff4451abdd883041d01f9e4@thread.v2"
    result = get_last_message_of_chat(access_token, chat_id)

    if result:
        # Extract the message content and remove HTML tags for simplicity
        message_content = result[0]['body']['content'].replace("<p>", "").replace("</p>", "").strip()
        
        # Convert the message to lowercase and check if it starts with 'phatgpt'
        if message_content.lower().startswith('phatgpt'):
            # Modify the message or create a new message as needed
            new_message = message_content.lower()  # Example modification, adjust as needed
            
            # Post the modified or new message back to the chat
            post_message_to_chat(access_token, chat_id, new_message)
            print("Message posted successfully.")
        else:
            print("The last message does not start with 'phatgpt'.")
    else:
        print("No messages found in the chat.")

if __name__ == "__main__":
    main()