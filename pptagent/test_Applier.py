from dotenv import load_dotenv
import anthropic
import os
import json
import urllib.request
import urllib.error
import traceback
import time
load_dotenv()

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

def _generate_code(api_key, type, before, after, slide_num, total_content):
    # API 호출을 위한 프롬프트 구성
    prompt = f"""
    Generate Python code to modify an active PowerPoint presentation based on the provided JSON task data.

    The code should:
    1. Find the slide specified by page number: {slide_num}
    2. Content reference: {total_content}
    3. Target to change: {type}, current content: {before}
    4. New content to apply: {after}
    5. Generate ONLY executable code that will directly modify the PowerPoint.

    CRITICAL REQUIREMENTS:
    - DO NOT create a new PowerPoint application - use the existing one
    - The code should NOT be written as a complete program with imports - it will be executed in an environment where PowerPoint is already open
    - Focus ONLY on finding and modifying the specified content
    - For text changes, use both shape.Name and TextFrame.TextRange.Text to identify the correct element
    - Always print detailed status messages for debugging
    - Make sure to explicitly apply any changes (e.g., shape.TextFrame.TextRange.Text = new_text)
    - Print each step of the process (e.g., "Found slide X", "Found shape Y", "Updated content from Z to W")

    Note that the code will run in a context where these variables are already available:
    - ppt_application: The PowerPoint application instance
    - active_presentation: The currently open presentation

    The code must be direct, practical and focused solely on making the specific change requested.
    """

    # Claude API 호출
    code = _call_claude_api(prompt, api_key)
    
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
    def __init__(self, api_key=os.environ.get('ANTHROPIC_API_KEY')):
        self.api_key = api_key
    
    def __call__(self, processed_json):
        # PowerPoint 설정 코드를 별도 함수로 실행
        try:
            import win32com.client
            import win32com.client.dynamic
            
            # PowerPoint 애플리케이션 가져오기
            print("Connecting to PowerPoint...")
            ppt_application = win32com.client.Dispatch("PowerPoint.Application")
            ppt_application.Visible = True  # PowerPoint 창 표시
            
            # 현재 활성화된 프레젠테이션 가져오기
            active_presentation = ppt_application.ActivePresentation
            print(f"Connected to PowerPoint. Presentation: {active_presentation.Name}")
            
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
                    code = _generate_code(self.api_key, type, before, after, slide_num, total_content)
                    
                    # 디버깅을 위해 전체 코드 출력
                    print("\n--- Generated Code ---")
                    print(code)
                    print("--- End of Code ---\n")
                    
                    # 코드 실행
                    try:
                        print("Executing code...")
                        
                        # 로컬 네임스페이스 생성 (각 반복마다 새로 생성)
                        # 글로벌 변수를 여기에 복사
                        local_vars = globals_dict.copy()
                        
                        # 코드 실행
                        exec(code, globals(), local_vars)
                        
                        # 오류 발생 여부 확인
                        if local_vars.get('error_occurred', False):
                            print("Error occurred during code execution")
                            success = False
                        else:
                            print("Code executed successfully")
                            
                            # 명시적으로 변경사항 저장
                            try:
                                active_presentation.Save()
                                print("Changes saved to presentation")
                            except Exception as save_error:
                                print(f"Error saving presentation: {save_error}")
                                success = False
                        
                        # 잠시 대기하여 PowerPoint가 업데이트될 시간 제공
                        time.sleep(0.5)
                        
                    except Exception as e:
                        print(f"Error executing code: {str(e)}")
                        print(f"Error type: {type(e).__name__}")
                        print(traceback.format_exc())
                        success = False
            
            return success
            
        except Exception as e:
            print(f"Error setting up PowerPoint: {str(e)}")
            print(traceback.format_exc())
            return False