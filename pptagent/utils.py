from llm_api import llm_request_with_retries
from prompt import PLAN_PROMPT, PLAN_INPUT_EX, PLAN_OUTPUT_EX
class Planner:
    def __init__(self):
        self.system_prompt = PLAN_PROMPT
        
        self.example_input = PLAN_INPUT_EX
        self.example_output = PLAN_OUTPUT_EX
    
    def __call__(self, user_input: str, model_name ="gemini-1.5-flash") -> dict:
        # Construct the prompt for the LLM
        prompt = f"""
        {self.system_prompt}
        
        Example Input:
        {self.example_input}
        
        Example Output:
        {self.example_output}
        
        Now, please create a plan for the following request:
        {user_input}
        """
        
        # Request plan from LLM
        response = llm_request_with_retries(
            model_name=model_name,
            request=prompt,
            num_retries=4
        )
        print(response)
        # The response should be a JSON string, but let's handle errors safely
        try:
            import json
            plan = json.loads(response)
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
    class VBAgenerator:
        pass
    pass

class Processor:
    pass

class Applier:
    class VBAgenerator:
        pass
    pass

class Reporter:
    pass

class SharedLogMemory:
    pass