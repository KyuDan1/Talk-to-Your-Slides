import google.generativeai as genai
import os
from PIL import Image
import numpy as np
import re
from openai import OpenAI
from loguru import logger
import yaml
import requests
from io import BytesIO
import base64
import random


credentials = yaml.safe_load(open("credentials.yml"))


SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
]


def convert_openai_to_gemini(openai_request):
    gemini_messages = []

    for message in openai_request["messages"]:
        parts = []
        for content in message["content"]:
            if isinstance(content, str):
                parts.append(content)
            elif content["type"] == "text":
                parts.append(content["text"])
            elif content["type"] == "image_url":
                image_url = content["image_url"]["url"]
                if image_url.startswith("data:image"):
                    # Extract base64 string and decode
                    base64_str = image_url.split(",")[1]
                    img_data = base64.b64decode(base64_str)
                    img = Image.open(BytesIO(img_data))
                else:
                    # Load the image from the URL
                    response = requests.get(image_url)
                    img = Image.open(BytesIO(response.content))
                parts.append(img)
        
        gemini_messages.append({"role": message["role"].replace("assistant", "model"), "parts": parts})
    
    # Extract parameters
    temperature = openai_request.get("temperature", 0.4)
    max_tokens = openai_request.get("max_tokens", 8192)
    stop = openai_request.get("stop", [])
    top_p = openai_request.get("top_p", 1.0)
    
    # Ensure stop is a list
    if isinstance(stop, str):
        stop = [stop]
    
    # Create the Gemini request
    gemini_request = {
        "contents": gemini_messages,
        "generation_config": {
            "temperature": temperature,
            "max_output_tokens": max_tokens,
            "stop_sequences": stop,
            "top_p": top_p,
        }
    }
    
    return gemini_request


def send_gemini_request(
    gemini_request,
    model,
    api_key=None
):
    if api_key is None:
        assert model in credentials, f"Model {model} not found in credentials"
        # Randomly select an API key if multiple are provided
        if "round-robin" in credentials[model]:
            num_keys = len(credentials[model]["round-robin"])
            rand_idx = random.randint(0, num_keys - 1)
            credential = credentials[model]["round-robin"][rand_idx]
        else:
            credential = credentials[model]
        api_key = credential["api_key"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model)
    
    logger.debug(
        f"API key: {'*' * (len(api_key) - 4)}{api_key[-4:]}"
    )
    
    response = model.generate_content(
        **gemini_request,
        safety_settings=SAFETY_SETTINGS,
    )
    return response.text
