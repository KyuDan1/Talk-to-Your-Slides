from dotenv import load_dotenv
from utils import _call_gpt_api
import anthropic
import os
import json
import urllib.request
import urllib.error
import traceback
import time
load_dotenv()
import openai
from openai import OpenAI
import pythoncom

    

def _call_claude_api(prompt, api_key):
    # Prepare API request
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    # API request data
    data = {
        "model": "claude-3-7-sonnet-20250219",
        "messages": [
            {
                "role": "user",
                "content": f"{prompt}"
            }
        ],
        "max_tokens": 4000
    }
    
    # Request configuration
    req = urllib.request.Request(url)
    for key, value in headers.items():
        req.add_header(key, value)
    
    # Encode JSON data
    json_data = json.dumps(data).encode('utf-8')
    
    # API call
    with urllib.request.urlopen(req, json_data) as response:
        response_text = response.read().decode('utf-8')
        result = json.loads(response_text)
        
        print("API response received")
        return result["content"][0]["text"]

def _generate_code(model, api_key, type, before, after, slide_num, total_content):
    # Prepare prompt for API call
    prompt = f"""
Generate Python code modify an active PowerPoint presentation based on the provided JSON task data. The code should:
0. Find activate powerpoint app with ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
   active_presentation = ppt_app.ActivePresentation
1. Find the slide specified by page number: {slide_num}
2. Content reference: {total_content}
3. Target to change: {type}, current content: {before}
4. New content to apply: {after}
5. Generate ONLY executable code that will directly modify the PowerPoint.

CRITICAL REQUIREMENTS:
- DO NOT create a new PowerPoint application - use the existing one
- Please check if the slide number you want to work on exists and proceed with the work. The index starts with 1.
- The code should NOT be written as a complete program with imports - it will be executed in an environment where PowerPoint is already open
- Focus on finding and modifying the specified content
- For text changes, use both shape.Name and TextFrame.TextRange.Text to identify the correct element
- Make sure to explicitly apply any changes (e.g., shape.TextFrame.TextRange.Text = new_text)
- Do not write print function or comments.
- You can write at slide note with slide.NotesPage
    ```python
    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text
    ```
Note that the code will run in a context where these variables are already available:
- ppt_application: The PowerPoint application instance
- active_presentation: The currently open presentation

IMPORTANT: In PowerPoint, color codes use BGR format (not RGB). For example, RGB(255,0,0) for red should be specified as RGB(0,0,255) in the code. Always convert any color references accordingly.

If you want to modify the formatting, refer to the following code for modification:
if text_frame.HasText:
    text_range = text_frame.TextRange
    # Find text
    found_range = text_range.Find(text_to_highlight)
    while found_range:
        found_any = True
        found_range.Font.Bold = True # Bold
        found_range.Font.Color.RGB = 255 # Example color (RED in BGR format - 0,0,255)
        found_range.Font.Size = found_range.Font.Size * 1.2 # Increase font size by 20%
        start_pos = found_range.Start + len(text_to_highlight)
        found_range = text_range.Find(text_to_highlight, start_pos)

Do not use any "**" to make bold. It won't be applied on powerpoint.
- You can add or split a page with 'presentation = ppt_app.Presentations.Add()'.

The code must be direct, practical and focused solely on making the specific change requested. Ensure all color references use the BGR format for proper appearance in PowerPoint.
    """

    # Call Claude API
    if model == "claude-3.7-sonnet":
        code = _call_claude_api(prompt, api_key)
    elif "gpt" in model:
        code = _call_gpt_api(prompt, api_key, model)
    else:
        "yes"
        code = _call_gpt_api(prompt, api_key, model)
    # Preprocess code
    code = code.strip()
    if code.startswith("```python"):
        code = code[len("```python"):].strip()
    if code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()
    
    # Check if required variables are defined in the code
    if "error_occurred = False" not in code:
        code = "error_occurred = False\n\n" + code
    
    return code
class test_Applier:
    def __init__(self, model = 'claude-3.7-sonnet', api_key=os.environ.get('ANTHROPIC_API_KEY'), retry = 3):
        self.api_key = api_key
        self.model = model
        self.retry = retry
    def __call__(self, processed_json):
        # Execute PowerPoint setup code in a separate function
        pythoncom.CoInitialize()
        try:
            import win32com.client
            import win32com.client.dynamic
            
            # Get PowerPoint application
            ppt_application = win32com.client.Dispatch("PowerPoint.Application")
            ppt_application.Visible = True  # Show PowerPoint window
            
            # Get currently active presentation
            active_presentation = ppt_application.ActivePresentation
            
            # Set global variables - explicitly define here
            globals_dict = {
                'ppt_application': ppt_application,
                'active_presentation': active_presentation,
                'win32com': win32com
            }
            
            success = True
            
            for task_idx, task in enumerate(processed_json['tasks']):
                print(f"\nProcessing task {task_idx+1}:")
                print(task)
                
                try:
                    slide_num = task['page_number']
                except (KeyError, TypeError):
                    try:
                        slide_num = task['page number']
                    except (KeyError, TypeError):
                        # Handle the case where neither key exists
                        print(f"Error: Could not find page number in task: {task}")
                        slide_num = None  # Or some default value
                
                total_content = task['contents']
                
                edit_target_type = task['edit target']
                edit_target_content = task['edit target content']
                content_after_edit = task['content after edit']
                
                for type, before, after in zip(edit_target_type, edit_target_content, content_after_edit):
                    if before == after:
                        print("No change needed - content is identical")
                        continue
                    
                    print(f"\nChanging {type}: '{before}' → '{after}'")
                    
                    # Generate code
                    code = _generate_code(self.model, self.api_key, type, before, after, slide_num, total_content)
                    
                    # Print full code for debugging
                    print("====Generated Code====")
                    print(code)
                    print("====End of Code=====")
                    
                    # Add retry logic for code execution
                    task_success = False
                    
                    for attempt in range(self.retry + 1):  # Initial attempt + retry count
                        if attempt > 0:
                            print(f"Retry attempt {attempt}/{self.retry}...")
                        
                        try:
                            print("Executing code...")
                            
                            # Create local namespace (new for each iteration)
                            local_vars = globals_dict.copy()
                            
                            # Execute code
                            exec(code, globals(), local_vars)
                            
                            # Check if error occurred
                            if local_vars.get('error_occurred', False):
                                print("Error occurred during code execution")
                                # Proceed to next retry (since error occurred)
                            else:
                                print("Code executed successfully")
                                task_success = True
                                break  # Break retry loop on success
                            
                            # Wait briefly to allow PowerPoint to update
                            time.sleep(0.5)
                            
                        except Exception as e:
                            print(f"Error executing code: {str(e)}")
                            print(f"Error type: {type(e).__name__}")
                            print(traceback.format_exc())
                            # Proceed to next retry (since exception occurred)
                    
                    # Set overall success status to False if all retries failed
                    if not task_success:
                        print(f"Failed to execute code after {self.retry} retries. Moving to next task.")
                        success = False
            
            return success
            
        except Exception as e:
            print(f"Error setting up PowerPoint: {str(e)}")
            print(traceback.format_exc())
            return False