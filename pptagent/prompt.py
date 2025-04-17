from utils import get_simple_powerpoint_info

PLAN_PROMPT = f"""You are a planning assistant for PowerPoint modifications.
Your job is to create a clear, step-by-step plan for modifying a PowerPoint presentation based on the user's request.
present ppt state: {get_simple_powerpoint_info()}
Break down complex requests into actionable tasks that can be executed by a PowerPoint automation system.
Focus on identifying:
1. Specific slides to modify
2. Specific object elements to add, remove, or change
3. The logical sequence of operations

Format your response as a JSON format with the following structure:
{{
    "understanding": "Brief summary of what the user wants to achieve",
    "tasks": [
        {{
            "page number": 1,
            "description": "Task description",
            "target": "Target slide or element",
            "action": "Action to perform",
        }},
        ...
    ],
}}

Below is the example question and example output.
input: Please translate the titles of slide 3 and slide 5 of the PPT into English.
output:
{{
    "understanding": "English translation of slide titles",
    "tasks": [
        {{
            "page number": 3,
            "description": "Translate the title of the slide",
            "target": "Title section",
            "action": "Translate to English",
        }},
        {{
            "page number": 5,
            "description": "Translate the title of the slide",
            "target": "Title section",
            "action": "Translate to English",
        }}
    ],
}}
"""



ACCESS_TO_VBA_PROJECT = """
PowerPoint의 VBA 프로젝트 액세스 보안 설정이 활성화되어 있어야 함


PowerPoint 보안 설정 확인:

PowerPoint를 열고 File > Options > Trust Center > Trust Center Settings > Macro Settings로 이동
"Trust access to the VBA project object model" 옵션을 체크해야 합니다
"""

PARSER_PROMPT = """


"""


VBA_PROMPT = """


"""

def create_process_prompt(page_number, description, action, contents):
    prompt = f"""Information about slide {page_number}:
- Task description: {description}
- Action type: {action}
- Slide contents: {contents}

Please analyze what and how to modify based on the contents of the above slide.
Identify the appropriate edit targets in the slide and determine what actions need to be taken for each target.
Consider the type of content (such as titles, content placeholders, text boxes, etc.) when making your assessment.
After identifying the targets, perform the necessary edits on each target to create the improved content. The edits might involve translation, formatting changes, content restructuring, or other modifications depending on what would improve the slide.
Please provide your analysis and recommendations in JSON format with three key arrays:

'edit target type': specifying the type of each content element you're modifying
'edit target content': listing the original content for each target
'content after edit': providing the modified version of each target content

Please edit only the {page_number}.
There might be cases where different texts within a single text box need to be edited separately.
For example, you could change only the English text to blue color.
Consider this and create a detailed plan.Each list should be the same length, and the edit targets type, edit target contents, contents after edit should correspond one-to-one.
"""
    return prompt