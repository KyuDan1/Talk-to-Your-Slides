# baseline1 Instruction-to-Code
from llm_api import llm_request_with_retries
from classes import Parser
BASELINE_PROMPT = f"""Create a Python code that can edit PowerPoint presentations by executing the following command:
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

def baseline1 (model_name="gemini-1.5-flash", user_instruction=None):
    

    # 전체 다 파싱할지? 어떻게 파싱하는게 baseline인가..
    parsed_data = Parser()
    
    response = llm_request_with_retries(model_name=model_name,
                            request=  PARSED_PROMPT
                                    + parsed_data
                                    + BASELINE_PROMPT
                                    + user_instruction
                                    + FORMAT_PROMPT)
    #response 에서 python 코드만 파싱하고
    response_code = parsing_python(response)
    exec(response_code)
    return response_code