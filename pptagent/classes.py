from llm_api import llm_request_with_retries
from prompt import PLAN_PROMPT, PARSER_PROMPT, VBA_PROMPT, create_process_prompt
from utils import parse_active_slide_objects
import json
import time
class Planner:
    def __init__(self):
        self.system_prompt = PLAN_PROMPT
    
    def __call__(self, user_input: str, model_name ="gemini-1.5-flash") -> dict:
        # Construct the prompt for the LLM
        prompt = f"""
        {self.system_prompt}
        
        Now, please create a plan for the following request:
        {user_input}
        """
        
        # Request plan from LLM
        response = llm_request_with_retries(
            model_name=model_name,
            request=prompt,
            num_retries=4
        )
        #print(response)
        # The response should be a JSON string, but let's handle errors safely
        try:
            import json
            json_str = response.strip().replace("```json", "").replace("```", "").strip()
            last_brace = json_str.rfind('}')
            if last_brace != -1:
                json_str = json_str[:last_brace+1]
            plan = json.loads(json_str)
            return plan
        except json.JSONDecodeError:
            # Fallback - return a basic structure with the raw response
            return {
                "understanding": "Failed to parse LLM response into proper JSON format",
                "tasks": [
                    {
                        "id": 1,
                        "description": "Review and manually interpret the plan",
                        "target": "N/A",
                        "action": "manual_review",
                        "details": response
                    }
                ],
                "requires_parsing": False,
                "requires_processing": False,
                "additional_notes": "LLM response was not in valid JSON format"
            }

class Parser:
    def __init__(self, json_data):
        self.json_data = json_data
        self.tasks = json_data.get("tasks", [])
    def process(self):
        """
        Process the tasks in the JSON data and add contents for each page number.
        
        Returns:
            dict: The updated JSON data with contents added
        """
        for task in self.tasks:
            page_number = task.get("page number")
            if page_number:
                # Parse objects for this slide
                slide_contents = parse_active_slide_objects(page_number)
                
                # Add the contents to the task
                task["contents"] = slide_contents
        
        return self.json_data
    

class Processor:
    def __init__(self, json_data, model_name="gemini-1.5-flash"):

        self.json_data = json_data
        # 'tasks' 키를 사용하는 것으로 보입니다 (원래 코드에서는 processed_datas였지만 입력 예시와 맞춤)
        self.tasks = json_data.get("tasks", [])
        self.model_name = model_name

    def process(self):
        for task in self.tasks:
            page_number = task.get("page number")
            description = task.get("description", "")
            action = task.get("action", "")
            contents = task.get("contents", "")
            
            if page_number:
                # LLM에게 보낼 프롬프트 생성
                prompt = create_process_prompt(page_number, description, action, contents)
                
                # LLM 요청 보내기
                response = llm_request_with_retries(
                    model_name=self.model_name,
                    request=prompt,
                    num_retries=4
                )
                
                # LLM 응답 파싱
                # print(response)
                json_str = response.strip().replace("```json", "").replace("```", "").strip()
                last_brace = json_str.rfind('}')
                if last_brace != -1:
                    json_str = json_str[:last_brace+1]
                processed_result = json.loads(json_str)

                
                # 결과를 작업에 추가
                task["edit target type"] = processed_result["edit target type"]
                task["edit target content"] = processed_result["edit target content"]
                task["content after edit"] = processed_result["content after edit"]
        
        return self.json_data

class Applier:
    class VBAgenerator:
        pass
    pass

class Reporter:
    pass

class SharedLogMemory:
    pass

    
class VBAgenerator:
    def __init__(self):
        self.system_prompt = VBA_PROMPT
   
    def __call__(self, user_input: str, model_name ="gemini-1.5-flash") -> dict:
        # Construct the prompt for the LLM
        prompt = f"""
        {self.system_prompt}
        
        Now, please create a python code using win32com library:
        {user_input}
        """
        
        # Request plan from LLM
        response = llm_request_with_retries(
            model_name=model_name,
            request=prompt,
            num_retries=4
        )
        print(response)