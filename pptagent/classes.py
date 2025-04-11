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

# 텍스트 수정: 기존 기능을 유지하면서 텍스트를 이름으로 찾지 못할 경우 내용으로 검색하는 기능 추가
# 텍스트 서식 변경: 글꼴, 크기, 볼드, 이탤릭, 밑줄, 색상, 정렬 등 변경
# 도형 색상 변경: 도형 채우기 색상 변경
# 테두리선 변경: 선 색상, 두께, 스타일, 투명도 변경
# 이미지 교체: 기존 이미지를 새 이미지로 교체
# 표 수정: 표 내용 및 서식 변경
# 차트 수정: 차트 데이터 및 제목 변경

# 데이터를 모은다면 Applier 데이터를 모아야 함.
class Applier:
    def __init__(self):
        """Initialize the Applier class"""
        pass
    
    def __call__(self, processed_json):
        """Apply the changes to PowerPoint based on the processed JSON
        
        Args:
            processed_json (dict): JSON containing editing tasks
            
        Returns:
            bool: True if modifications were successful, False otherwise
        """
        try:
            # Generate PowerPoint modification code
            code = self._generate_code(processed_json)
            
            # Execute the code in a separate namespace
            local_vars = {}
            exec(code, globals(), local_vars)
            
            # Check if any errors occurred during execution
            if local_vars.get('error_occurred', False):
                return False
            
            return True
        except Exception as e:
            print(f"Error applying changes: {e}")
            return False
    
    def _generate_code(self, processed_json):
        """Generate Python code to modify the PowerPoint presentation
        
        Args:
            processed_json (dict): JSON containing editing tasks
            
        Returns:
            str: Python code as a string
        """
        code_lines = [
            "import win32com.client",
            "import re",
            "from collections import defaultdict",
            "",
            "error_occurred = False",
            "",
            "# Helper functions for different edit operations",
            "def modify_text(shape, new_text):",
            "    if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:",
            "        shape.TextFrame.TextRange.Text = new_text",
            "        return True",
            "    return False",
            "",
            "def modify_text_formatting(shape, format_dict):",
            "    if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:",
            "        text_range = shape.TextFrame.TextRange",
            "        if 'font' in format_dict:",
            "            text_range.Font.Name = format_dict['font']",
            "        if 'size' in format_dict:",
            "            text_range.Font.Size = format_dict['size']",
            "        if 'bold' in format_dict:",
            "            text_range.Font.Bold = -1 if format_dict['bold'] else 0",
            "        if 'italic' in format_dict:",
            "            text_range.Font.Italic = -1 if format_dict['italic'] else 0",
            "        if 'underline' in format_dict:",
            "            text_range.Font.Underline = -1 if format_dict['underline'] else 0",
            "        if 'color' in format_dict:",
            "            # Expects RGB tuple",
            "            r, g, b = format_dict['color']",
            "            text_range.Font.Color.RGB = r + (g * 256) + (b * 256 * 256)",
            "        if 'alignment' in format_dict:",
            "            align_map = {'left': 1, 'center': 2, 'right': 3, 'justify': 4}",
            "            text_range.ParagraphFormat.Alignment = align_map.get(format_dict['alignment'].lower(), 1)",
            "        return True",
            "    return False",
            "",
            "def modify_shape_fill(shape, color):",
            "    if hasattr(shape, 'Fill'):",
            "        r, g, b = color",
            "        shape.Fill.ForeColor.RGB = r + (g * 256) + (b * 256 * 256)",
            "        return True",
            "    return False",
            "",
            "def modify_shape_line(shape, line_props):",
            "    if hasattr(shape, 'Line'):",
            "        if 'color' in line_props:",
            "            r, g, b = line_props['color']",
            "            shape.Line.ForeColor.RGB = r + (g * 256) + (b * 256 * 256)",
            "        if 'weight' in line_props:",
            "            shape.Line.Weight = line_props['weight']",
            "        if 'style' in line_props:",
            "            style_map = {'solid': 1, 'dash': 2, 'dot': 3, 'dash-dot': 4, 'dash-dot-dot': 5}",
            "            shape.Line.DashStyle = style_map.get(line_props['style'].lower(), 1)",
            "        if 'transparency' in line_props:",
            "            shape.Line.Transparency = line_props['transparency']",
            "        return True",
            "    return False",
            "",
            "def replace_image(shape, image_path):",
            "    try:",
            "        # Get shape position and size",
            "        left = shape.Left",
            "        top = shape.Top",
            "        width = shape.Width",
            "        height = shape.Height",
            "        slide = shape.Parent",
            "        # Delete the old shape",
            "        old_shape_name = shape.Name",
            "        shape.Delete()",
            "        # Add new picture",
            "        new_shape = slide.Shapes.AddPicture(",
            "            FileName=image_path,",
            "            LinkToFile=False,",
            "            SaveWithDocument=True,",
            "            Left=left,",
            "            Top=top,",
            "            Width=width,",
            "            Height=height",
            "        )",
            "        # Set the same name as the old shape",
            "        new_shape.Name = old_shape_name",
            "        return True",
            "    except Exception as e:",
            "        print(f'Error replacing image: {e}')",
            "        return False",
            "",
            "def modify_table(shape, table_data):",
            "    if hasattr(shape, 'Table'):",
            "        table = shape.Table",
            "        if 'data' in table_data:",
            "            data = table_data['data']",
            "            for i, row in enumerate(data):",
            "                for j, cell_text in enumerate(row):",
            "                    if i < table.Rows.Count and j < table.Columns.Count:",
            "                        table.Cell(i+1, j+1).Shape.TextFrame.TextRange.Text = str(cell_text)",
            "        if 'formatting' in table_data and isinstance(table_data['formatting'], dict):",
            "            formatting = table_data['formatting']",
            "            # Apply cell-specific formatting",
            "            if 'cells' in formatting:",
            "                for cell_info in formatting['cells']:",
            "                    row, col = cell_info['position']",
            "                    if row < table.Rows.Count and col < table.Columns.Count:",
            "                        cell = table.Cell(row+1, col+1).Shape",
            "                        if 'fill_color' in cell_info:",
            "                            r, g, b = cell_info['fill_color']",
            "                            cell.Fill.ForeColor.RGB = r + (g * 256) + (b * 256 * 256)",
            "                        if 'text_format' in cell_info:",
            "                            modify_text_formatting(cell, cell_info['text_format'])",
            "        return True",
            "    return False",
            "",
            "def modify_chart(shape, chart_data):",
            "    if hasattr(shape, 'Chart'):",
            "        chart = shape.Chart",
            "        # Excel needs to be used to modify chart data",
            "        xl = win32com.client.Dispatch('Excel.Application')",
            "        xl.Visible = False",
            "        try:",
            "            # Get chart data workbook",
            "            wb = chart.ChartData.Workbook",
            "            # Activate workbook",
            "            xl_wb = xl.Workbooks.Open(wb.FullName)",
            "            ws = xl_wb.Worksheets(1)",
            "            # Modify data",
            "            if 'series_data' in chart_data:",
            "                for series_idx, series_values in chart_data['series_data'].items():",
            "                    # Excel is 1-based, series_idx should be 0-based in input",
            "                    row = int(series_idx) + 2  # +2 because row 1 is usually headers",
            "                    for col_idx, value in enumerate(series_values):",
            "                        ws.Cells(row, col_idx + 2).Value = value  # +2 as col A is often labels",
            "            # Update chart title if provided",
            "            if 'title' in chart_data and chart_data['title']:",
            "                chart.ChartTitle.Text = chart_data['title']",
            "            # Save changes",
            "            xl_wb.Save()",
            "            xl_wb.Close()",
            "            # Refresh the chart",
            "            chart.Refresh()",
            "            return True",
            "        except Exception as e:",
            "            print(f'Error modifying chart: {e}')",
            "            return False",
            "        finally:",
            "            xl.Quit()",
            "    return False",
            "",
            "try:",
            "    # Initialize PowerPoint application",
            "    ppt_app = win32com.client.Dispatch('PowerPoint.Application')",
            "    ppt_app.Visible = True",
            "",
            "    # Get the active presentation",
            "    try:",
            "        presentation = ppt_app.ActivePresentation",
            "    except:",
            "        print('Error: No active PowerPoint presentation found.')",
            "        error_occurred = True",
            "        raise Exception('No active PowerPoint presentation found.')",
            ""
        ]
        
        # Process each task in the processed JSON
        for task in processed_json['tasks']:
            slide_number = task['page number']
            
            code_lines.extend([
                f"    # Processing slide {slide_number}",
                f"    slide = presentation.Slides({slide_number})",
                ""
            ])
            
            # Get action type from task
            action = task.get('action', '').lower()
            
            # Default behavior - text edits
            if 'edit target type' in task and 'edit target content' in task and 'content after edit' in task:
                edit_types = task['edit target type']
                original_contents = task['edit target content']
                new_contents = task['content after edit']
                
                # Iterate through each potential edit
                for i, (edit_type, original, new) in enumerate(zip(edit_types, original_contents, new_contents)):
                    # Skip if no change needed
                    if original == new:
                        code_lines.append(f"    # No change needed for '{edit_type}' with content '{original}'")
                        continue
                    
                    # Handle special characters in strings
                    new_escaped = new.replace("'", "\\'")
                    
                    code_lines.extend([
                        f"    # Change '{edit_type}' from '{original}' to '{new_escaped}'",
                        f"    shape_found = False",
                        f"    for shape in slide.Shapes:",
                        f"        if shape.Name == '{edit_type}':",
                        f"            shape_found = True",
                        f"            success = modify_text(shape, '{new_escaped}')",
                        f"            if success:",
                        f"                print(f'Changed {edit_type} text to: {new_escaped}')",
                        f"            else:",
                        f"                print(f'Shape {edit_type} does not support text modification')",
                        f"                error_occurred = True",
                        f"            break",
                        f"",
                        f"    if not shape_found:",
                        f"        print(f'Shape with name \"{edit_type}\" not found on slide {slide_number}')",
                        f"        # Try alternative search by examining all shapes with text",
                        f"        text_found = False",
                        f"        for shape in slide.Shapes:",
                        f"            if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:",
                        f"                if shape.TextFrame.TextRange.Text.strip() == '{original}':",
                        f"                    text_found = True",
                        f"                    success = modify_text(shape, '{new_escaped}')",
                        f"                    if success:",
                        f"                        print(f'Changed shape with matching text to: {new_escaped}')",
                        f"                    else:",
                        f"                        print(f'Error modifying text in shape with matching content')",
                        f"                        error_occurred = True",
                        f"                    break",
                        f"        if not text_found:",
                        f"            print(f'Could not find any shape with text \"{original}\" on slide {slide_number}')",
                        f"            error_occurred = True",
                        f""
                    ])
            
            # Handle formatting changes if specified
            if 'formatting' in task:
                formatting = task['formatting']
                if isinstance(formatting, list) and len(formatting) > 0:
                    for fmt_item in formatting:
                        target = fmt_item.get('target', '')
                        format_type = fmt_item.get('type', '').lower()
                        properties = fmt_item.get('properties', {})
                        
                        if format_type == 'text':
                            code_lines.extend([
                                f"    # Apply text formatting to '{target}'",
                                f"    for shape in slide.Shapes:",
                                f"        if shape.Name == '{target}':",
                                f"            success = modify_text_formatting(shape, {properties})",
                                f"            if success:",
                                f"                print(f'Applied text formatting to {target}')",
                                f"            else:",
                                f"                print(f'Failed to apply text formatting to {target}')",
                                f"                error_occurred = True",
                                f"            break",
                                f""
                            ])
                        elif format_type == 'fill':
                            code_lines.extend([
                                f"    # Apply fill color to '{target}'",
                                f"    for shape in slide.Shapes:",
                                f"        if shape.Name == '{target}':",
                                f"            success = modify_shape_fill(shape, {properties.get('color', (255, 255, 255))})",
                                f"            if success:",
                                f"                print(f'Applied fill color to {target}')",
                                f"            else:",
                                f"                print(f'Failed to apply fill color to {target}')",
                                f"                error_occurred = True",
                                f"            break",
                                f""
                            ])
                        elif format_type == 'line':
                            code_lines.extend([
                                f"    # Apply line properties to '{target}'",
                                f"    for shape in slide.Shapes:",
                                f"        if shape.Name == '{target}':",
                                f"            success = modify_shape_line(shape, {properties})",
                                f"            if success:",
                                f"                print(f'Applied line properties to {target}')",
                                f"            else:",
                                f"                print(f'Failed to apply line properties to {target}')",
                                f"                error_occurred = True",
                                f"            break",
                                f""
                            ])
            
            # Handle image replacements
            if 'images' in task:
                images = task['images']
                if isinstance(images, list) and len(images) > 0:
                    for img_item in images:
                        target = img_item.get('target', '')
                        new_path = img_item.get('path', '')
                        
                        code_lines.extend([
                            f"    # Replace image '{target}' with '{new_path}'",
                            f"    for shape in slide.Shapes:",
                            f"        if shape.Name == '{target}':",
                            f"            success = replace_image(shape, r'{new_path}')",
                            f"            if success:",
                            f"                print(f'Replaced image {target} with {new_path}')",
                            f"            else:",
                            f"                print(f'Failed to replace image {target}')",
                            f"                error_occurred = True",
                            f"            break",
                            f"    else:",
                            f"        print(f'Image shape \"{target}\" not found on slide {slide_number}')",
                            f"        error_occurred = True",
                            f""
                        ])
            
            # Handle table modifications
            if 'tables' in task:
                tables = task['tables']
                if isinstance(tables, list) and len(tables) > 0:
                    for table_item in tables:
                        target = table_item.get('target', '')
                        table_data = table_item.get('data', {})
                        
                        code_lines.extend([
                            f"    # Modify table '{target}'",
                            f"    for shape in slide.Shapes:",
                            f"        if shape.Name == '{target}':",
                            f"            success = modify_table(shape, {table_data})",
                            f"            if success:",
                            f"                print(f'Modified table {target}')",
                            f"            else:",
                            f"                print(f'Failed to modify table {target}')",
                            f"                error_occurred = True",
                            f"            break",
                            f"    else:",
                            f"        print(f'Table \"{target}\" not found on slide {slide_number}')",
                            f"        error_occurred = True",
                            f""
                        ])
            
            # Handle chart modifications
            if 'charts' in task:
                charts = task['charts']
                if isinstance(charts, list) and len(charts) > 0:
                    for chart_item in charts:
                        target = chart_item.get('target', '')
                        chart_data = chart_item.get('data', {})
                        
                        code_lines.extend([
                            f"    # Modify chart '{target}'",
                            f"    for shape in slide.Shapes:",
                            f"        if shape.Name == '{target}':",
                            f"            success = modify_chart(shape, {chart_data})",
                            f"            if success:",
                            f"                print(f'Modified chart {target}')",
                            f"            else:",
                            f"                print(f'Failed to modify chart {target}')",
                            f"                error_occurred = True",
                            f"            break",
                            f"    else:",
                            f"        print(f'Chart \"{target}\" not found on slide {slide_number}')",
                            f"        error_occurred = True",
                            f""
                        ])
        
        # Save the presentation and handle cleanup
        code_lines.extend([
            "    # Save the presentation",
            "    try:",
            "        presentation.Save()",
            "        print('PowerPoint presentation has been updated and saved successfully.')",
            "    except Exception as e:",
            "        print(f'Error saving presentation: {e}')",
            "        error_occurred = True",
            "",
            "except Exception as e:",
            "    print(f'Error: {e}')",
            "    error_occurred = True",
            "finally:",
            "    # Clean up (but don't close PowerPoint so user can see changes)",
            "    try:",
            "        if 'ppt_app' in locals():",
            "            pass  # Keep PowerPoint open for user to view changes",
            "    except Exception as e:",
            "        print(f'Error during cleanup: {e}')",
            "        error_occurred = True"
        ])
        
        return "\n".join(code_lines)

import os
import json
from datetime import datetime

class Reporter:
    def __init__(self):
        pass
        
    def __call__(self, processed_json, result):
        # Create a prompt for the LLM to summarize what was done
        prompt = self._create_prompt(processed_json, result)
        
        # Request summary from LLM
        summary = llm_request_with_retries(
            model_name="gemini-1.5-flash",
            request=prompt,
            num_retries=4
        )
        
        # Print and return the summary
        # print("Report to user:")
        # print(summary)
        return summary
        
    def _create_prompt(self, processed_json, result):
        # Create a prompt that instructs the LLM to summarize the actions taken
        prompt = f"""
        Summarize the following processing information and result into a concise report for the user:
        
        Processed information: {json.dumps(processed_json, indent=2, ensure_ascii=False)}
        Result: {"Success" if result else "Failure"}
        
        Your summary should:
        1. Explain what task was attempted in simple terms
        2. Mention the specific actions that were taken
        3. State whether the task was completed successfully
        4. If the task failed, briefly explain why
        5. Be concise and user-friendly
        6. If the task involved translation, include both the original text and the translated text
        
        Summary:
        """
        return prompt

class SharedLogMemory:
    def __init__(self, log_dir="logs"):
        self.log_dir = log_dir
        # Create log directory if it doesn't exist
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
    
    def __call__(self, user_input, plan_json, processed_json, result):
        # Create a timestamp for the log
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create log record
        log_record = {
            "timestamp": timestamp,
            "user_input": user_input,
            "plan": plan_json,
            "processed": processed_json,
            "result": result
        }
        
        # Save the log record as a JSON file
        log_file_path = os.path.join(self.log_dir, f"log_{timestamp}.json")
        with open(log_file_path, "w", encoding="utf-8") as f:
            json.dump(log_record, f, indent=2, ensure_ascii=False)
        
        # Also save as a text file for easier human reading
        log_text_path = os.path.join(self.log_dir, f"log_{timestamp}.txt")
        with open(log_text_path, "w", encoding="utf-8") as f:
            f.write(f"Timestamp: {timestamp}\n\n")
            f.write(f"User Input: {user_input}\n\n")
            f.write(f"Plan:\n{json.dumps(plan_json, indent=2, ensure_ascii=False)}\n\n")
            f.write(f"Processed:\n{json.dumps(processed_json, indent=2, ensure_ascii=False)}\n\n")
            f.write(f"Result:\n{json.dumps(result, indent=2, ensure_ascii=False)}\n\n")
        
        # Return the log record as memories
        memories = log_record
        return memories

    
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