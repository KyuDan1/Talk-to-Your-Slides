# baseline1 Instruction-to-Code
from llm_api import llm_request_with_retries, llm_request_with_retries_gemini
from utils import _call_gpt_api
from dotenv import load_dotenv
load_dotenv()
from classes import Parser
import os
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
GEMINI_API_KEY = os.environ.get('GEMINI')
BASELINE_PROMPT = f"""Create a Python code with win32com library that can edit PowerPoint presentations by executing the following command:
                    """
FORMAT_PROMPT = """IMPORTANT: Your response must contain ONLY valid Python code wrapped in triple backticks with the 'python' language tag. Follow this exact format:
```python
# Your Python code here
# Include proper comments, imports, and function definitions
# No explanations or text outside this code block
```
"""
PARSED_PROMPT = "The following is information parsed from a PPT slide."

def parsing_python(text):
    # Look for Python code blocks with markdown formatting
    import re
    
    # Pattern to match code blocks with explicit Python language tag
    python_pattern = r"```python\s*([\s\S]*?)\s*```"
    python_matches = re.findall(python_pattern, text)
    
    # If we found Python code blocks specifically tagged with "python"
    if python_matches:
        return python_matches[0].strip()
    
    # If no specific Python blocks found, try to match any code blocks
    generic_pattern = r"```\s*([\s\S]*?)\s*```"
    generic_matches = re.findall(generic_pattern, text)
    
    if generic_matches:
        return generic_matches[0].strip()
    
    # If no code blocks found, check if the entire response might be code
    # (no markdown formatting)
    if "def " in text or "import " in text:
        # Try to extract what looks like code - this is a simple heuristic
        lines = text.split('\n')
        code_lines = []
        for line in lines:
            # Skip lines that look like explanations or markdown
            if line.startswith('#') or not line.strip():
                code_lines.append(line)
            elif re.match(r'^[a-zA-Z0-9_\s()\[\]{}=+\-*/.<>,:\'\"]+$', line):
                code_lines.append(line)
        
        return '\n'.join(code_lines).strip()
    
    # If nothing else worked, return empty string
    return "" 



# # baseline2 Instruction + Parsed Info

import json
import traceback  # 에러 메시지를 얻기 위해 필요

def baseline1(model_name="gemini-1.5-flash", user_instruction=None):
    from classes import Parser
    from llm_api import llm_request_with_retries_gemini
    
    # 현재 열린 ppt의 전체 페이지 수 가져오기
    parser = Parser(baseline=True)
    parsed_json = parser.process()
    
    parsed_json_str = json.dumps(parsed_json, ensure_ascii=False)
    
    if 'gpt' in model_name:
        response, input_token, output_token, price = _call_gpt_api(
        model=model_name,
        prompt=PARSED_PROMPT + str(parsed_json_str) + BASELINE_PROMPT + user_instruction + FORMAT_PROMPT,
        api_key=OPENAI_API_KEY
    )
    else:
        response, input_token, output_token, price = llm_request_with_retries_gemini(
            model_name=model_name,
            prompt_content=PARSED_PROMPT + str(parsed_json_str) + BASELINE_PROMPT + user_instruction + FORMAT_PROMPT
        )
    
    # Python 코드만 파싱
    response_code = parsing_python(response)
    
    # Try to execute the code
    execution_success = False
    execution_error = None
    try:
        exec(response_code)
        execution_success = True
    except Exception as e:
        execution_success = False
        execution_error = traceback.format_exc()  # 에러 전체 traceback 캡처

    return execution_success, input_token, output_token, price, response_code, execution_error


#response_code, input_token, output_token, price, code, err = baseline1(model_name="gpt-4.1-mini", user_instruction="translate slide number 1 in English.")
#print(response_code, input_token, output_token, price, code, err)