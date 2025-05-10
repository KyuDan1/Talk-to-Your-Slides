import os
from loguru import logger
from openai import OpenAI
import yaml
import time
import re

import random
from dotenv import load_dotenv
load_dotenv()
from gemini_api import convert_openai_to_gemini, send_gemini_request
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
GEMINI_API_KEY = os.environ.get('GEMINI')

credentials = yaml.safe_load(open("./credentials.yml"))


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

import google.generativeai as genai
import time
import os
import logging
import json # 원래 코드의 JSONL 로깅 부분을 위해 남겨둠

GEMINI_PRICING = {
    # Vertex AI 및 새 API 명명 규칙에 따른 모델 이름 예시
    
    "gemini-1.5-flash": {
        "input_per_million_tokens": 0.075, # 1M 토큰당 (일반적인 사용 사례, >128k 컨텍스트는 다를 수 있음)
        "output_per_million_tokens": 0.3, # 1M 토큰당
    },
    
    "gemini-2.5-flash-preview-04-17": {
        "input_per_million_tokens": 0.15,
        "output_per_million_tokens": 0.6,
    }
}


def llm_request_with_retries_gemini(model_name: str,
                             prompt_content,
                             num_retries: int = 4,
                             ) -> tuple[str, int, int, float]:
    from google import genai
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    # Count input tokens before sending the request
    input_tokens = 0
    try:
        token_count = client.models.count_tokens(
            model=model_name, contents=prompt_content
        )
        input_tokens = token_count.total_tokens
    except Exception as e:
        logger.warning(f"Error counting input tokens: {e}")
    
    for attempt in range(num_retries):
        try:
            request = create_openai_request(content=prompt_content)
            
            #response = client.models.generate_content(model = model_name, contents=request)
            response , output_token_count = send_openai_request(request, model_name)
            
            # Get the final response text
            final_response_text = response
            
            # Get output tokens from response metadata
            output_tokens = 0
            output_tokens = output_token_count
            # 'candidates_token_count' 뒤의 숫자만 매칭
            # m = re.search(r'candidates?_token_count\s*:\s*(\d+)', metadata_str)
            # if m:
            #     output_tokens = int(m.group(1))
            # else:
            #     raise ValueError(f"'candidates_token_count'를 usage_metadata에서 찾을 수 없습니다: {metadata_str}")
            #output_tokens = response.usage_metadata['candidated_token_count']
            # try:
            #     # Assuming response has usage_metadata attribute
            #     if hasattr(response, 'usage_metadata'):
            #         output_tokens = response.usage_metadata.candidates_token_count
            #     elif isinstance(response, dict) and 'usage_metadata' in response:
            #         output_tokens = response['usage_metadata']['candidates_token_count']
            # except Exception as e:
            #     logger.warning(f"Error extracting output tokens: {e}")
            
            # Calculate cost based on pricing
            total_cost = 0.0
            if model_name in GEMINI_PRICING:
                input_cost = (input_tokens / 1_000_000) * GEMINI_PRICING[model_name]["input_per_million_tokens"]
                output_cost = (output_tokens / 1_000_000) * GEMINI_PRICING[model_name]["output_per_million_tokens"]
                total_cost = input_cost + output_cost
            
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
        return "", 0, 0, 0.0
    
    return final_response_text, input_tokens, output_tokens, total_cost