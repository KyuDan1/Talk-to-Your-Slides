import os
from loguru import logger
from openai import OpenAI
import yaml
import time
import random
from gemini_api import convert_openai_to_gemini, send_gemini_request


credentials = yaml.safe_load(open("credentials.yml"))


def create_openai_request(
    content:str,
    temperature=0.4,
    stop=["’’’’", " – –", "<|endoftext|>", "<|eot_id|>"],
    
    
):
    
    
    messages = [
    {
        "role": "user",
        "content": content
        + "\n\n"
        
    }
                ]
    
    return {
        "messages": messages,
        "temperature": temperature,
        "stop": stop
    }


def openai_client(
    model,
    api_key=None,
    base_url="https://api.openai.com/v1"
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
        if "base_url" in credential:
            base_url = credential["base_url"]
    client = OpenAI(api_key=api_key, base_url=base_url)
    
    logger.debug(
        f"API key: ****{api_key[-4:]}, endpoint: {base_url}"
    )
    
    return client

def send_openai_request(
    openai_request,
    model,
    api_key=None,
    base_url="https://api.openai.com/v1"
):
    if "gemini" in model:
        return send_gemini_request(
            convert_openai_to_gemini(openai_request),
            model,
            api_key=api_key
        )
    client = openai_client(model, api_key=api_key, base_url=base_url)
    
    response = client.chat.completions.create(
        model=model, **openai_request
    )
    return response.choices[0].message.content

def llm_request_with_retries(model_name: str,
                             request,
                             num_retries: int = 4,
                             )->str:
    for attempt in range(num_retries):
        try:
            request = create_openai_request(content=request)
            response = send_openai_request(request, model_name)
            # Write the result to jsonl
            """with open(jsonl_fn, 'a') as f:
                json.dump({'custom_id': custom_id, 'request': request, 'response': response}, f)
                f.write('\n')"""
            # If successful, break the retry loop
            break
        except Exception as e:
            if "503" in str(e):  # Server not up yet, sleep until the server is up again
                while True:
                    logger.debug("503 error, sleep 30 seconds")
                    time.sleep(30)
                    try:
                        response = send_openai_request(request, model_name)
                        break
                    except Exception as e:
                        if "503" not in str(e):
                            break
            else:
                logger.error(e)
                # If an exception occurs, wait and then retry
                wait_time = 2 ** (attempt + 3)
                logger.debug(f"Attempt {attempt + 1} failed. Waiting for {wait_time} seconds before retrying...")
                time.sleep(wait_time)
                continue
    else:
        logger.error(f"Failed to process after {num_retries} attempts")
    
    return response

# 사용법법
"""response = llm_request_with_retries(model_name="gemini-1.5-flash",
                         request="hey there?")
print(response)"""