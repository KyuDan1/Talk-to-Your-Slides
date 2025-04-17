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

    

def _call_claude_api(prompt, api_key):
    # API 요청 준비
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    # API 요청 데이터
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
    
    # 요청 설정
    req = urllib.request.Request(url)
    for key, value in headers.items():
        req.add_header(key, value)
    
    # JSON 데이터 인코딩
    json_data = json.dumps(data).encode('utf-8')
    
    # API 호출
    with urllib.request.urlopen(req, json_data) as response:
        response_text = response.read().decode('utf-8')
        result = json.loads(response_text)
        
        print("API 응답 수신 완료")
        return result["content"][0]["text"]

def _generate_code(model, api_key, type, before, after, slide_num, total_content):
    # API 호출을 위한 프롬프트 구성
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
- Please check if the slide number you want to work on exists and proceed with the work.
- The code should NOT be written as a complete program with imports - it will be executed in an environment where PowerPoint is already open
- Focus on finding and modifying the specified content
- For text changes, use both shape.Name and TextFrame.TextRange.Text to identify the correct element
- Make sure to explicitly apply any changes (e.g., shape.TextFrame.TextRange.Text = new_text)
- Print each step of the process (e.g., "Found slide X", "Found shape Y", "Updated content from Z to W")
- Do not write print function or comments.

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

    # Claude API 호출
    if model == "claude-3.7-sonnet":
        code = _call_claude_api(prompt, api_key)
    elif "gpt" in model:
        code = _call_gpt_api(prompt, api_key, model)
    else:
        "yes"
        code = _call_gpt_api(prompt, api_key, model)
    # 코드 전처리
    code = code.strip()
    if code.startswith("```python"):
        code = code[len("```python"):].strip()
    if code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()
    
    # 코드에 필요한 변수가 정의되어 있는지 확인
    if "error_occurred = False" not in code:
        code = "error_occurred = False\n\n" + code
    
    return code

class test_Applier:
    def __init__(self, model = 'claude-3.7-sonnet', api_key=os.environ.get('ANTHROPIC_API_KEY'), retry = 3):
        self.api_key = api_key
        self.model = model
        self.retry = retry
    def __call__(self, processed_json):
        # PowerPoint 설정 코드를 별도 함수로 실행
        try:
            import win32com.client
            import win32com.client.dynamic
            
            # PowerPoint 애플리케이션 가져오기
            #print("Connecting to PowerPoint...")
            ppt_application = win32com.client.Dispatch("PowerPoint.Application")
            ppt_application.Visible = True  # PowerPoint 창 표시
            
            # 현재 활성화된 프레젠테이션 가져오기
            active_presentation = ppt_application.ActivePresentation
            #print(f"Connected to PowerPoint. Presentation: {active_presentation.Name}")
            
            # 글로벌 변수 설정 - 여기서 명시적으로 정의
            globals_dict = {
                'ppt_application': ppt_application,
                'active_presentation': active_presentation,
                'win32com': win32com
            }
            
            success = True
            
            for task_idx, task in enumerate(processed_json['tasks']):
                print(f"\nProcessing task {task_idx+1}:")
                print(task)
                
                slide_num = task['page number']
                total_content = task['contents']
                
                edit_target_type = task['edit target type']
                edit_target_content = task['edit target content']
                content_after_edit = task['content after edit']
                
                for type, before, after in zip(edit_target_type, edit_target_content, content_after_edit):
                    if before == after:
                        print("No change needed - content is identical")
                        continue
                    
                    print(f"\nChanging {type}: '{before}' → '{after}'")
                    
                    # 코드 생성
                    code = _generate_code(self.model, self.api_key, type, before, after, slide_num, total_content)
                    
                    # 디버깅을 위해 전체 코드 출력
                    print("====Generated Code====")
                    print(code)
                    print("====End of Code=====")
                    
                    # 코드 실행에 retry 로직 추가
                    task_success = False
                    
                    for attempt in range(self.retry + 1):  # 초기 시도 + retry 횟수만큼 시도
                        if attempt > 0:
                            print(f"Retry attempt {attempt}/{self.retry}...")
                        
                        try:
                            print("Executing code...")
                            
                            # 로컬 네임스페이스 생성 (각 반복마다 새로 생성)
                            local_vars = globals_dict.copy()
                            
                            # 코드 실행
                            exec(code, globals(), local_vars)
                            
                            # 오류 발생 여부 확인
                            if local_vars.get('error_occurred', False):
                                print("Error occurred during code execution")
                                # 다음 재시도로 진행 (에러가 발생했으므로)
                            else:
                                print("Code executed successfully")
                                task_success = True
                                break  # 성공했으므로 재시도 루프 종료
                            
                            # 잠시 대기하여 PowerPoint가 업데이트될 시간 제공
                            time.sleep(0.5)
                            
                        except Exception as e:
                            print(f"Error executing code: {str(e)}")
                            print(f"Error type: {type(e).__name__}")
                            print(traceback.format_exc())
                            # 다음 재시도로 진행 (예외가 발생했으므로)
                    
                    # 모든 재시도 후에도 실패한 경우 전체 성공 상태를 False로 설정
                    if not task_success:
                        print(f"Failed to execute code after {self.retry} retries. Moving to next task.")
                        success = False
            
            return success
            
        except Exception as e:
            print(f"Error setting up PowerPoint: {str(e)}")
            print(traceback.format_exc())
            return False