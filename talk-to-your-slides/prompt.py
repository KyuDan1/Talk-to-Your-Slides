from utils import get_simple_powerpoint_info

PLAN_PROMPT = fPLAN_PROMPT = f"""You are a planning assistant for PowerPoint modifications.
Your job is to create a detailed, specific, step-by-step plan for modifying a PowerPoint presentation based on the user's request.
present ppt state: {get_simple_powerpoint_info()}
Break down complex requests into highly specific actionable tasks that can be executed by a PowerPoint automation system.
Focus on identifying:
1. Specific slides to modify (by page number)
2. Specific sections within slides (title, body, notes, headers, footers, etc.)
3. Specific object elements to add, remove, or change (text boxes, images, shapes, charts, tables, etc.)
4. Precise formatting changes (font, size, color, alignment, etc.)
5. The logical sequence of operations with clear dependencies

Format your response as a JSON format with the following structure:
{{
    "understanding": "Detailed summary of what the user wants to achieve",
    "tasks": [
        {{
            "page number": 1,
            "description": "Specific task description",
            "target": "Precise target location (e.g., 'Title section of slide 1', 'Notes section of slide 3', 'Second bullet point in body text', 'Chart in bottom right')",
            "action": "Specific action with all necessary details",
            "contents": {{
                "additional details required for the action"
            }}
        }},
        ...
    ],
}}

Below is the example question and example output.
input: Please translate the titles of slide 3 and slide 5 of the PPT into English.
output:
{{
    "understanding": "English translation of slide titles on slides 3 and 5",
    "tasks": [
        {{
            "page number": 3,
            "description": "Translate the title text of slide 3",
            "target": "Title section of slide 3",
            "action": "Translate to English",
            "contents": {{
                "source_language": "auto-detect",
                "preserve_formatting": true
            }}
        }},
        {{
            "page number": 5,
            "description": "Translate the title text of slide 5",
            "target": "Title section of slide 5",
            "action": "Translate to English",
            "contents": {{
                "source_language": "auto-detect",
                "preserve_formatting": true
            }}
        }}
    ],
}}

Be extremely specific and detailed in your task descriptions and targeting. For example:
- "Notes section of slide 3" instead of just "slide 3"
- "Third bullet point in body text of slide 7" instead of "body text"
- "Blue rectangular shape in the top-right corner of slide 2" instead of "shape"
- "Chart title in the data visualization on slide 4" instead of "chart"

Response in JSON format.
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

'edit target': specifying the type and exact ID of each content element you're modifying
'edit target content': listing the original content for each target
'content after edit': providing the modified version of each target content
Be careful not to edit elements that the user has not requested.

Please edit only the {page_number}.
There might be cases where different texts within a single text box need to be edited separately.
For example, you could change only the English text to blue color.
Consider this and create a detailed plan.Each list should be the same length, and the edit targets type, edit target contents, contents after edit should correspond one-to-one.
Give me in JSON format.
"""
    return prompt