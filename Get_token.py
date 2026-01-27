import requests as rq
import json
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


def get_token():
    """Get authentication token from the custom endpoint"""
    # Try to get credentials from environment variables first, then from credencial.json
    email = os.getenv("token_signin_email") 
    password = os.getenv("token_signin_password") 
    url = os.getenv("token_signin_url")
     

    if not email or not password:
        print("❌ Email and password not found in environment variables or credencial.json")
        return None


    headers = {
        'Content-Type': 'application/json'
    }

    data = {
        'email': email,
        'password': password
    }
    
    response = rq.post(url, json=data, headers=headers)

    # Assuming the response contains a token
    if response.status_code == 200:
        token_data = response.json()
        token = token_data.get('token')  # Adjust based on actual response structure
        if token:
            return token
        else:
            print("❌ Token not found in response")
            return None
    else:
        print(f"❌ Failed to get token: {response.status_code}, {response.text}")
        return None
    
    
    
# print(get_token())