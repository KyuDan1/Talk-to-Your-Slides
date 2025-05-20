import win32com.client
import pywintypes
import openai
from openai import OpenAI
import re

def parse_llm_response(response):
    """
    Function to reliably parse JSON from LLM response
    """
    import json
    import re
    
    try:
        # Handle empty response
        if not response:
            print("Response is empty.")
            return None
            
        # Initial cleanup
        json_str = response.strip()
        
        # Remove markdown code blocks
        json_str = re.sub(r'```(?:json)?', '', json_str)
        json_str = json_str.replace('```', '').strip()
        
        # Check if it's valid JSON from the start
        try:
            return json.loads(json_str)
        except:
            pass  # Direct parsing failed, continue
        
        # Find JSON object boundaries
        start_idx = json_str.find('{')
        end_idx = json_str.rfind('}')
        
        if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
            json_str = json_str[start_idx:end_idx+1]
        else:
            # Check array format
            start_idx = json_str.find('[')
            end_idx = json_str.rfind(']')
            if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
                json_str = json_str[start_idx:end_idx+1]
            else:
                print("No valid JSON structure found.")
                return None
        
        # Remove control characters (except tab, newline, carriage return)
        json_str = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]+', '', json_str)
        
        # Fix common JSON errors
        # Remove trailing commas
        json_str = re.sub(r',\s*}', '}', json_str)
        json_str = re.sub(r',\s*]', ']', json_str)
        
        # Check if keys are quoted
        json_str = re.sub(r'([{,])\s*([a-zA-Z0-9_]+)\s*:', r'\1"\2":', json_str)
        
        # Convert single quotes to double quotes
        json_str = re.sub(r"'([^']*)'", r'"\1"', json_str)
        
        # Handle special escape characters
        json_str = json_str.replace('\\n', '\\\\n').replace('\\r', '\\\\r').replace('\\t', '\\\\t')
        
        # Handle nested quotes
        json_str = re.sub(r'(?<!\\)"(?=.*":)', '\\"', json_str)
        
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"JSON parsing error position: {e.pos}, message: {e.msg}")
            
            # Print around error position
            error_pos = e.pos
            start = max(0, error_pos - 20)
            end = min(len(json_str), error_pos + 20)
            print(f"Problem area: ...{json_str[start:error_pos]}|HERE|{json_str[error_pos:end]}...")
            
            # Last attempt: remove problematic character and retry
            if error_pos < len(json_str):
                fixed_str = json_str[:error_pos] + json_str[error_pos+1:]
                try:
                    return json.loads(fixed_str)
                except:
                    pass
            
            # If failed, try more aggressive cleanup
            # Remove non-ASCII characters
            clean_str = re.sub(r'[^\x20-\x7E]', '', json_str)
            try:
                return json.loads(clean_str)
            except:
                print("All JSON parsing attempts failed.")
                return None
                
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return None
    
def extract_content_after_edit(plan_json):
    result = []
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'content after edit' in task and isinstance(task['content after edit'], list):
                result.extend(task['content after edit'])
    
    return result

def extract_last_text_content(plan_json):
    last_text = ""
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'contents' in task:
                contents_str = task['contents']
                # Text content: 패턴을 모두 찾아서 리스트로 만듦
                text_contents = re.findall(r'Text content: (.*?)(?=\n\s+Font:|$)', contents_str, re.DOTALL)
                
                # 마지막 Text content: 내용을 반환 (없으면 빈 문자열)
                if text_contents:
                    last_text = text_contents[-1].strip()
    
    return last_text

def create_thinking_queue(plan_json):
    # thinking queue
    temp_tasks = []
    temp_actions = []
    
    print_data_ = ""

    for i in range(len(plan_json['tasks'])):
        temp_tasks.append(plan_json['tasks'][i]['target'])
        temp_actions.append(plan_json['tasks'][i]['action'])
    
    for i in range(len(temp_tasks)):
        print_data_ += f"• {temp_actions[i]} 작업을 '{temp_tasks[i]}'에 적용합니다.\n"
    
    return print_data_


def _call_gpt_api(prompt: str, api_key: str, model: str):
    # API 키 설정
    openai.api_key = api_key
    
    # 지원하는 모델 목록 검증
    allowed_models = ["gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano", "o4-mini"]
    if model not in allowed_models:
        raise ValueError(f"Model must be one of {allowed_models}")

    # 모델명 매핑 (== → =)
    if model == "gpt-4.1":
        model = "gpt-4.1-2025-04-14"
    elif model == "gpt-4.1-mini":
        model = "gpt-4.1-mini-2025-04-14"
    elif model == "gpt-4.1-nano":
        model = "gpt-4.1-nano-2025-04-14"
    else:
        model = model

    try:
        client = OpenAI(
            api_key = api_key,
        )
        response = client.responses.create(
            model=model,
            instructions="You are a coding assistant that editing powerpoint slides.",
            input=prompt,
        )
        
        return response.output_text

    except Exception as e:
        print(f"An error occurred: {e}")
        return None
def get_simple_powerpoint_info():
    """
    현재 열려있는 PowerPoint의 페이지 수와 파일 이름만 가져옵니다.
    """
    try:
        # PowerPoint 애플리케이션에 연결
        ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")
        
        # PowerPoint가 실행 중이고 열린 프레젠테이션이 있는지 확인
        if not ppt_app or not hasattr(ppt_app, 'ActivePresentation'):
            return "PowerPoint가 실행 중이 아니거나 열린 프레젠테이션이 없습니다."
        
        # 활성 프레젠테이션 가져오기
        presentation = ppt_app.ActivePresentation
        
        # 파일 이름과 페이지 수 가져오기
        file_name = presentation.Name
        slide_count = presentation.Slides.Count
        
        return {
            "파일 이름": file_name,
            "슬라이드 수": slide_count
        }
        
    except Exception as e:
        return f"오류 발생: {str(e)}"


def get_shape_type(shape_type):
    """Shape type number to string conversion"""
    shape_types = {
        1: "AutoShape", 
        2: "CallOut",
        3: "Chart",
        4: "Comment",
        5: "Freeform",
        6: "Group",
        7: "EmbeddedOLEObject",
        8: "FormControl",
        9: "Line",
        10: "LinkedOLEObject",
        11: "LinkedPicture",
        12: "OLEControl",
        13: "Picture",
        14: "Placeholder",
        15: "MediaObject", 
        16: "TextEffect",
        17: "TextBox",
        18: "Table",
        19: "SmartArt",
        20: "WebVideo",
        21: "ContentApp"
    }
    return shape_types.get(shape_type, f"Unknown Type ({shape_type})")

def get_placeholder_type(placeholder_type):
    """Placeholder type number to string conversion"""
    placeholder_types = {
        1: "Title",
        2: "Body",
        3: "CenterTitle",
        4: "SubTitle",
        5: "VerticalTitle",
        6: "VerticalBody",
        7: "Object",
        8: "Chart",
        9: "Table",
        10: "ClipArt",
        11: "OrgChart",
        12: "Media",
        13: "VerticalObject",
        14: "Picture",
        15: "Slide Number",
        16: "Header",
        17: "Footer",
        18: "Date",
        19: "VerticalTitle2",
        20: "VerticalBody2" 
    }
    return placeholder_types.get(placeholder_type, f"Unknown Placeholder ({placeholder_type})")

def parse_text_frame(text_frame):
    """Text frame information parsing"""
    result = ""
    try:
        if text_frame.HasText:
            result += f"\n   - Text: {text_frame.TextRange.Text}"
            result += f"\n   - Paragraphs Count: {text_frame.TextRange.Paragraphs().Count}"
            result += f"\n   - Text Alignment: {text_frame.TextRange.ParagraphFormat.Alignment}"
            result += f"\n   - Font Name: {text_frame.TextRange.Font.Name}"
            result += f"\n   - Font Size: {text_frame.TextRange.Font.Size}"
            result += f"\n   - Font Color: RGB({text_frame.TextRange.Font.Color.RGB & 0xFF}, {(text_frame.TextRange.Font.Color.RGB >> 8) & 0xFF}, {(text_frame.TextRange.Font.Color.RGB >> 16) & 0xFF})"
            
            # Check for hyperlinks
            try:
                hyperlinks = text_frame.TextRange.ActionSettings(1).Hyperlink
                if hasattr(hyperlinks, 'Address') and hyperlinks.Address:
                    result += f"\n   - Hyperlink: {hyperlinks.Address}"
            except:
                pass
    except Exception as e:
        result += f"\n   - Text Frame Error: {e}"
    return result

def parse_table(table):
    """Table information parsing"""
    result = ""
    try:
        rows = table.Rows.Count
        cols = table.Columns.Count
        result += f"\n   - Table Dimensions: {rows}x{cols} (Rows x Columns)"
        
        # Table style and properties
        result += f"\n   - First Row: {table.FirstRow}"
        result += f"\n   - Last Row: {table.LastRow}"
        result += f"\n   - First Column: {table.FirstCol}"
        result += f"\n   - Last Column: {table.LastCol}"
        
        # Get text content from all cells
        # Since the amount of text can be large,
        # only sample the first row and column
        if rows > 0 and cols > 0:
            result += "\n   - Sample Cell Contents:"
            max_sample = min(3, rows)
            for r in range(1, max_sample + 1):
                for c in range(1, min(3, cols) + 1):
                    try:
                        cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                        result += f"\n     Cell({r},{c}): {cell_text[:30]}{'...' if len(cell_text) > 30 else ''}"
                    except:
                        result += f"\n     Cell({r},{c}): [Error reading cell]"
    except Exception as e:
        result += f"\n   - Table Parsing Error: {e}"
    return result

def parse_chart(chart):
    """Chart information parsing"""
    result = ""
    try:
        chart_type = chart.ChartType
        chart_types = {
            -4100: "xlColumnClustered",
            -4101: "xlColumnStacked",
            -4170: "xlBarClustered",
            -4102: "xlLineStacked",
            73: "xlPie",
            # Add more chart types as needed
        }
        chart_type_name = chart_types.get(chart_type, f"Unknown Chart Type ({chart_type})")
        
        result += f"\n   - Chart Type: {chart_type_name}"
        result += f"\n   - Has Legend: {chart.HasLegend}"
        result += f"\n   - Has Title: {chart.HasTitle}"
        
        if chart.HasTitle:
            result += f"\n   - Chart Title: {chart.ChartTitle.Text}"
        
        # Chart data series information
        try:
            series_count = chart.SeriesCollection().Count
            result += f"\n   - Series Count: {series_count}"
            
            # Series name example
            if series_count > 0:
                result += "\n   - Series Names:"
                for i in range(1, min(series_count + 1, 5)):  # Show up to 4 series
                    try:
                        series_name = chart.SeriesCollection(i).Name
                        result += f"\n     Series {i}: {series_name}"
                    except:
                        result += f"\n     Series {i}: [Error reading series]"
        except Exception as se:
            result += f"\n   - Series Error: {se}"
    except Exception as e:
        result += f"\n   - Chart Parsing Error: {e}"
    return result

def parse_group_shapes(group_shape):
    """Group internal Shape information parsing"""
    result = ""
    try:
        shapes_count = group_shape.GroupItems.Count
        result += f"\n   - Group Items Count: {shapes_count}"
        
        # Parse all group items to avoid too much output
        # Only parse the first few items briefly
        max_items = min(3, shapes_count)
        for i in range(1, max_items + 1):
            try:
                sub_shape = group_shape.GroupItems.Item(i)
                result += f"\n   - Group Item {i}: {sub_shape.Name} (Type: {get_shape_type(sub_shape.Type)})"
            except:
                result += f"\n   - Group Item {i}: [Error reading item]"
    except Exception as e:
        result += f"\n   - Group Parsing Error: {e}"
    return result

def parse_picture(picture):
    """Image information parsing"""
    result = ""
    try:
        result += f"\n   - Picture Type: {picture.Type}"
        
        # Image size ratio
        if hasattr(picture, 'ScaleHeight') and hasattr(picture, 'ScaleWidth'):
            result += f"\n   - Scale: Width {picture.ScaleWidth}%, Height {picture.ScaleHeight}%"
        
        # Image quality and compression information
        if hasattr(picture, 'PictureFormat'):
            pf = picture.PictureFormat
            if hasattr(pf, 'Brightness'):
                result += f"\n   - Brightness: {pf.Brightness}"
            if hasattr(pf, 'Contrast'):
                result += f"\n   - Contrast: {pf.Contrast}"
            if hasattr(pf, 'Crop'):
                crop = pf.Crop
                result += f"\n   - Crop: Left={crop.ShapeLeft}, Top={crop.ShapeTop}, Right={crop.ShapeWidth}, Bottom={crop.ShapeHeight}"
    except Exception as e:
        result += f"\n   - Picture Parsing Error: {e}"
    return result

def parse_placeholder_details(placeholder):
    """Placeholder detailed information parsing"""
    result = ""
    try:
        ph_type = placeholder.PlaceholderFormat.Type
        result += f"\n   - Placeholder Type: {get_placeholder_type(ph_type)} (ID: {ph_type})"
        result += f"\n   - Placeholder ID: {placeholder.Id}"
        result += f"\n   - Placeholder Index: {placeholder.PlaceholderFormat.Index}"
        
        # Additional Placeholder properties
        if hasattr(placeholder.PlaceholderFormat, 'ContainedType'):
            result += f"\n   - Contained Type: {placeholder.PlaceholderFormat.ContainedType}"
    except Exception as e:
        result += f"\n   - Placeholder Parsing Error: {e}"
    return result

def parse_shape_details(shape):
    """
    Parse details by shape type
    """
    details = {
        "type": shape.Type,
        "name": shape.Name,
        "text": "",
        "has_text": False,
        "has_table": False,
        "table_data": None,
        "has_chart": False,
        "chart_type": None,
        "has_image": False,
        "image_type": None
    }
    
    # Parse text content
    if shape.HasTextFrame:
        text_frame = shape.TextFrame
        if text_frame.HasText:
            details["text"] = text_frame.TextRange.Text
            details["has_text"] = True
    
    # Parse table content
    if shape.HasTable:
        details["has_table"] = True
        table = shape.Table
        table_data = []
        for row in range(1, table.Rows.Count + 1):
            row_data = []
            for col in range(1, table.Columns.Count + 1):
                cell = table.Cell(row, col)
                if cell.Shape.HasTextFrame:
                    row_data.append(cell.Shape.TextFrame.TextRange.Text)
                else:
                    row_data.append("")
            table_data.append(row_data)
        details["table_data"] = table_data
    
    # Parse chart content
    if shape.HasChart:
        details["has_chart"] = True
        chart = shape.Chart
        details["chart_type"] = chart.ChartType
    
    # Parse image content
    if shape.Type == 13:  # msoPicture
        details["has_image"] = True
        if shape.Fill.Type == -1:  # msoFillPicture
            details["image_type"] = "Picture"
    
    return details

def parse_slide_notes(slide):
    """
    Parse slide notes
    """
    if not slide.HasNotesPage:
        return ""
    
    notes_page = slide.NotesPage
    notes_text = ""
    
    # Find notes placeholder
    for shape in notes_page.Shapes:
        if shape.Type == 14:  # msoPlaceholder
            if shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                if shape.HasTextFrame:
                    text_frame = shape.TextFrame
                    if text_frame.HasText:
                        notes_text = text_frame.TextRange.Text
                        break
    
    return notes_text

def parse_slide_properties(slide):
    """Slide properties parsing"""
    result = "\n\n--- SLIDE PROPERTIES ---"
    try:
        # Slide basic information
        result += f"\nSlide ID: {slide.SlideID}"
        result += f"\nSlide Index: {slide.SlideIndex}"
        result += f"\nSlide Name: {slide.Name}"
        
        # Slide layout information
        if hasattr(slide, 'Layout'):
            result += f"\nSlide Layout Type: {slide.Layout.Type}"
            result += f"\nSlide Layout Name: {slide.Layout.Name}"
        
        # Slide background information
        if hasattr(slide, 'Background'):
            bg = slide.Background
            if hasattr(bg, 'Fill'):
                fill = bg.Fill
                fill_type = "None"
                if hasattr(fill, 'Type'):
                    fill_types = {1: "Solid", 2: "Pattern", 3: "Gradient", 4: "Texture", 5: "Picture"}
                    fill_type = fill_types.get(fill.Type, f"Unknown ({fill.Type})")
                result += f"\nBackground Fill Type: {fill_type}"
        
        # Slide transition effect
        if hasattr(slide, 'SlideShowTransition'):
            trans = slide.SlideShowTransition
            result += f"\nTransition: {trans.EntryEffect if hasattr(trans, 'EntryEffect') else 'None'}"
            result += f"\nAdvance Time: {trans.AdvanceTime if hasattr(trans, 'AdvanceTime') else 'Manual'} seconds"
            result += f"\nAdvance On Click: {'Yes' if trans.AdvanceOnClick else 'No'}"
            result += f"\nAdvance On Time: {'Yes' if trans.AdvanceOnTime else 'No'}"
        
    except Exception as e:
        result += f"\nError parsing slide properties: {e}"
    
    return result

def parse_active_slide_objects(slide_num:int=1):
    """Main function for parsing slide objects"""
    output = "" # Initialize string to store output
    
    try:
        # Connect to running PowerPoint instance
        ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
        # Get active presentation
        presentation = ppt.ActivePresentation
        
        # Check if there is an active presentation
        if not presentation:
            output += "No active presentation found."
            return output
        
        # Add presentation information
        output += f"Presentation: {presentation.Name}\n"
        output += f"Total Slides: {presentation.Slides.Count}\n"
        
        # Check slide range
        if slide_num > presentation.Slides.Count or slide_num < 1:
            output += f"Invalid slide number. Please provide a number between 1 and {presentation.Slides.Count}."
            return output
        
        # Access the specified slide
        slide = presentation.Slides(slide_num)
        
        # Slide properties parsing
        output += parse_slide_properties(slide)
        
        # Get the number of shapes in the slide
        shape_count = slide.Shapes.Count
        output += f"\n\n--- SLIDE OBJECTS ---\nFound {shape_count} objects in slide number {slide_num}."
        
        # Iterate through each shape
        for i in range(1, shape_count + 1):
            shape = slide.Shapes(i)
            output += f"\n\nObject {i}:"
            output += f"\n Name: {shape.Name}"
            output += f"\n Type: {get_shape_type(shape.Type)}"
            output += f"\n Position: Left={shape.Left}, Top={shape.Top}"
            output += f"\n Size: Width={shape.Width}, Height={shape.Height}"
            
            # Parse details based on shape type
            output += parse_shape_details(shape)
        
        # Slide notes parsing
        output += parse_slide_notes(slide)
        
        output += "\n\nParsing complete."
        
    except pywintypes.com_error as e:
        output += f"COM error: {e}"
    except Exception as e:
        output += f"Error: {e}"
    
    return output

# def parse_active_slide_objects(slide_num:int=1):
#     output = ""  # 출력을 저장할 문자열 초기화
#     try:
#         # Connect to running PowerPoint instance
#         ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
#         # Get active presentation
#         presentation = ppt.ActivePresentation
        
#         # Check if there is an active presentation
#         if not presentation:
#             output += "No active presentation found."
#             return output
        
#         # Access the first slide
#         slide = presentation.Slides(slide_num)
        
#         # Get the number of shapes in the slide
#         shape_count = slide.Shapes.Count
#         output += f"Found {shape_count} objects in the slide number {slide_num}."
        
#         # Iterate through each shape
#         for i in range(1, shape_count + 1):
#             shape = slide.Shapes(i)
#             output += f"\nObject {i}:"
#             output += f"\n  Name: {shape.Name}"
#             output += f"\n  Type: {get_shape_type(shape.Type)}"
#             output += f"\n  Position: Left={shape.Left}, Top={shape.Top}"
#             output += f"\n  Size: Width={shape.Width}, Height={shape.Height}"
            
#             # Parse details based on shape type
#             output += parse_shape_details(shape)
                
#         output += "\nParsing complete."
        
#     except pywintypes.com_error as e:
#         output += f"COM error: {e}"
#     except Exception as e:
#         output += f"Error: {e}"
    
#     return output

def get_shape_type(type_val):
    # Map shape type values to readable names
    # Official documentation: https://learn.microsoft.com/en-us/office/vba/api/office.msoshapetype
    shape_types = {
        1: "AutoShape",
        2: "Callout",
        3: "Chart",
        4: "Comment",
        5: "Freeform",
        6: "Group",
        7: "Embedded OLE Object",
        8: "Form Control",
        9: "Line",
        10: "Linked OLE Object",
        11: "Linked Picture",
        12: "OLE Control Object",
        13: "Picture",
        14: "Placeholder",
        15: "Text Effect",
        16: "Media",
        17: "Text Box",
        18: "Script Anchor",
        19: "Table",
        20: "Canvas",
        21: "Diagram",
        22: "Ink",
        23: "Ink Comment",
        24: "Smart Art",
        25: "Web Video",
        26: "Content App"
    }
    return shape_types.get(type_val, f"Unknown Type ({type_val})")

def extract_text_from_shape(shape, indent_level=1):
    """Function to extract text from all shape types"""
    output = ""
    indent = "  " * indent_level
    
    try:
        # Extract text from objects with text frame (TextBox, AutoShape, Placeholder, etc.)
        if hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
            text_frame = shape.TextFrame
            if hasattr(text_frame, 'HasText') and text_frame.HasText:
                text_range = text_frame.TextRange
                output += f"\n{indent}Text content: {text_range.Text}"
                
                # Extract text formatting information
                try:
                    output += f"\n{indent}Font: {text_range.Font.Name}, Size: {text_range.Font.Size}"
                    output += f"\n{indent}Bold: {text_range.Font.Bold}, Italic: {text_range.Font.Italic}"
                    
                    # Extract paragraph information
                    if hasattr(text_range, 'ParagraphFormat'):
                        para_format = text_range.ParagraphFormat
                        output += f"\n{indent}Alignment: {get_alignment_type(para_format.Alignment)}"
                        output += f"\n{indent}Line Spacing: {para_format.LineSpacing}"
                except:
                    output += f"\n{indent}Cannot retrieve all text formatting details"
        
        # TextFrame2 support (Office 2007 and above)
        elif hasattr(shape, 'HasTextFrame2') and shape.HasTextFrame2:
            text_frame = shape.TextFrame2
            if hasattr(text_frame, 'TextRange') and text_frame.TextRange.Text != "":
                text_range = text_frame.TextRange
                output += f"\n{indent}Text content (TextFrame2): {text_range.Text}"
                
                # Extract text formatting information
                try:
                    output += f"\n{indent}Font: {text_range.Font.Name}, Size: {text_range.Font.Size}"
                    output += f"\n{indent}Bold: {text_range.Font.Bold}, Italic: {text_range.Font.Italic}"
                except:
                    output += f"\n{indent}Cannot retrieve TextFrame2 formatting details"
        
        # Extract text from table cells
        elif shape.Type == 19:  # Table
            try:
                table = shape.Table
                rows = table.Rows.Count
                cols = table.Columns.Count
                
                output += f"\n{indent}Table text content:"
                for row in range(1, rows + 1):
                    for col in range(1, cols + 1):
                        cell = table.Cell(row, col)
                        if hasattr(cell, 'Shape') and hasattr(cell.Shape, 'TextFrame'):
                            text_frame = cell.Shape.TextFrame
                            if text_frame.HasText:
                                cell_text = text_frame.TextRange.Text
                                output += f"\n{indent}  Cell({row},{col}): {cell_text}"
            except:
                output += f"\n{indent}Cannot retrieve table text content"
                
        # Extract text from chart
        elif shape.Type == 3:  # Chart
            try:
                chart = shape.Chart
                if chart.HasTitle:
                    output += f"\n{indent}Chart Title: {chart.ChartTitle.Text}"
                    
                # Extract axis titles
                if hasattr(chart, 'Axes'):
                    for axis_type in range(1, 3):  # 1: Primary, 2: Secondary
                        for axis_group in range(1, 4):  # 1: X, 2: Y, 3: Z
                            try:
                                axis = chart.Axes(axis_group, axis_type)
                                if axis.HasTitle:
                                    output += f"\n{indent}Axis Title ({axis_group},{axis_type}): {axis.AxisTitle.Text}"
                            except:
                                pass
            except:
                output += f"\n{indent}Cannot retrieve chart text content"
                
        # Extract text from SmartArt
        elif shape.Type == 24:  # SmartArt
            try:
                if hasattr(shape, 'SmartArt'):
                    smart_art = shape.SmartArt
                    if hasattr(smart_art, 'AllNodes'):
                        nodes = smart_art.AllNodes
                        output += f"\n{indent}SmartArt text content:"
                        for i in range(1, nodes.Count + 1):
                            node = nodes.Item(i)
                            if hasattr(node, 'TextFrame2'):
                                text_frame = node.TextFrame2
                                if text_frame.TextRange.Text != "":
                                    output += f"\n{indent}  Node {i}: {text_frame.TextRange.Text}"
            except:
                output += f"\n{indent}Cannot retrieve SmartArt text content"
    
    except Exception as e:
        output += f"\n{indent}Text extraction error: {e}"
    
    return output

def get_alignment_type(alignment_val):
    # Convert paragraph alignment value to text
    alignment_types = {
        1: "Left",
        2: "Center",
        3: "Right",
        4: "Justify",
        5: "Distributed"
    }
    return alignment_types.get(alignment_val, f"Unknown Alignment ({alignment_val})")

def parse_group_shape(group_shape, indent_level=1):
    """Recursively parse all items within a group object"""
    output = ""  # Initialize string to store output
    try:
        indent = "  " * indent_level
        group_items_count = group_shape.GroupItems.Count
        output += f"{indent}Number of objects in group: {group_items_count}"
        
        # Iterate through each item in the group
        for j in range(1, group_items_count + 1):
            group_item = group_shape.GroupItems.Item(j)
            output += f"\n{indent}Object in group {j}:"
            output += f"\n{indent}  Name: {group_item.Name}"
            output += f"\n{indent}  Type: {get_shape_type(group_item.Type)}"
            output += f"\n{indent}  Position: Left={group_item.Left}, Top={group_item.Top}"
            output += f"\n{indent}  Size: Width={group_item.Width}, Height={group_item.Height}"
            
            # Extract text from the object
            output += extract_text_from_shape(group_item, indent_level + 1)
            
            # Process recursively if the item in the group is another group
            if group_item.Type == 6:  # Group
                output += parse_group_shape(group_item, indent_level + 1)
            else:
                # Parse regular shape details
                output += parse_shape_details(group_item, indent_level + 1)
                
    except Exception as e:
        output += f"\n{indent}Group object parsing error: {e}"
    
    return output

def parse_shape_details(shape, indent_level=1):
    # Extract specific details based on shape type
    output = ""  # Initialize string to store output
    indent = "  " * indent_level
    
    # Try extracting text from all shapes
    output += extract_text_from_shape(shape, indent_level)
    
    try:
        if shape.Type == 6:  # Group
            output += f"\n{indent}Group object found: {shape.Name}"
            output += parse_group_shape(shape, indent_level)
                
        elif shape.Type == 13:  # Picture
            output += f"\n{indent}Picture: {shape.Name}"
            try:
                output += f"\n{indent}Alternative Text: {shape.AlternativeText}"
            except:
                pass
                
        elif shape.Type == 3:  # Chart
            output += f"\n{indent}Chart: {shape.Name}"
            try:
                chart = shape.Chart
                output += f"\n{indent}Chart Type: {chart.ChartType}"
                output += f"\n{indent}Has Title: {chart.HasTitle}"
            except:
                output += f"\n{indent}Cannot retrieve all chart details"
                
        elif shape.Type == 19:  # Table
            output += f"\n{indent}Table: {shape.Name}"
            try:
                table = shape.Table
                output += f"\n{indent}Rows: {table.Rows.Count}, Columns: {table.Columns.Count}"
            except:
                output += f"\n{indent}Cannot retrieve all table details"
                
    except Exception as e:
        output += f"\n{indent}Shape details parsing error: {e}"
    
    return output

# output = parse_active_slide_objects()
# print(output)