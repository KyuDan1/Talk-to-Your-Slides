import win32com.client
import pywintypes
import openai
from openai import OpenAI
import re

def parse_llm_response(response):
    """
    LLM 응답에서 JSON을 안정적으로 파싱하는 함수
    """
    import json
    import re
    
    try:
        # 응답이 None이거나 빈 문자열인 경우 처리
        if not response:
            print("응답이 비어 있습니다.")
            return None
            
        # 초기 정리
        json_str = response.strip()
        
        # 마크다운 코드 블록 제거
        json_str = re.sub(r'```(?:json)?', '', json_str)
        json_str = json_str.replace('```', '').strip()
        
        # 처음부터 유효한 JSON인지 확인
        try:
            return json.loads(json_str)
        except:
            pass  # 직접 파싱 실패, 계속 진행
        
        # JSON 객체의 경계 찾기
        start_idx = json_str.find('{')
        end_idx = json_str.rfind('}')
        
        if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
            json_str = json_str[start_idx:end_idx+1]
        else:
            # 배열 형식 확인
            start_idx = json_str.find('[')
            end_idx = json_str.rfind(']')
            if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
                json_str = json_str[start_idx:end_idx+1]
            else:
                print("유효한 JSON 구조를 찾을 수 없습니다.")
                return None
        
        # 제어 문자 제거 (탭, 줄바꿈, 캐리지 리턴은 제외)
        json_str = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]+', '', json_str)
        
        # 일반적인 JSON 오류 수정
        # 후행 쉼표 제거
        json_str = re.sub(r',\s*}', '}', json_str)
        json_str = re.sub(r',\s*]', ']', json_str)
        
        # 키가 따옴표로 묶여 있는지 확인
        json_str = re.sub(r'([{,])\s*([a-zA-Z0-9_]+)\s*:', r'\1"\2":', json_str)
        
        # 작은따옴표를 큰따옴표로 변경
        json_str = re.sub(r"'([^']*)'", r'"\1"', json_str)
        
        # 특수 이스케이프 문자 처리
        json_str = json_str.replace('\\n', '\\\\n').replace('\\r', '\\\\r').replace('\\t', '\\\\t')
        
        # 중첩된 따옴표 처리
        json_str = re.sub(r'(?<!\\)"(?=.*":)', '\\"', json_str)
        
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"JSON 파싱 오류 위치: {e.pos}, 메시지: {e.msg}")
            
            # 문제가 있는 위치 주변 출력
            error_pos = e.pos
            start = max(0, error_pos - 20)
            end = min(len(json_str), error_pos + 20)
            print(f"문제 지점: ...{json_str[start:error_pos]}|HERE|{json_str[error_pos:end]}...")
            
            # 마지막 시도: 문제가 되는 문자 제거 후 재시도
            if error_pos < len(json_str):
                fixed_str = json_str[:error_pos] + json_str[error_pos+1:]
                try:
                    return json.loads(fixed_str)
                except:
                    pass
            
            # 실패한 경우 더 공격적인 정리 시도
            # ASCII가 아닌 문자 제거
            clean_str = re.sub(r'[^\x20-\x7E]', '', json_str)
            try:
                return json.loads(clean_str)
            except:
                print("모든 JSON 파싱 시도가 실패했습니다.")
                return None
                
    except Exception as e:
        print(f"예상치 못한 오류: {str(e)}")
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
    """Shape 유형 번호를 문자열로 변환"""
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
    """Placeholder 유형 번호를 문자열로 변환"""
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
    """텍스트 프레임 정보 파싱"""
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
    """테이블 정보 파싱"""
    result = ""
    try:
        rows = table.Rows.Count
        cols = table.Columns.Count
        result += f"\n   - Table Dimensions: {rows}x{cols} (Rows x Columns)"
        
        # 테이블 스타일 및 속성
        result += f"\n   - First Row: {table.FirstRow}"
        result += f"\n   - Last Row: {table.LastRow}"
        result += f"\n   - First Column: {table.FirstCol}"
        result += f"\n   - Last Column: {table.LastCol}"
        
        # 모든 셀의 텍스트 내용을 가져오기에는 양이 많을 수 있으므로
        # 첫 번째 행과 열의 일부만 샘플로 표시
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
    """차트 정보 파싱"""
    result = ""
    try:
        chart_type = chart.ChartType
        chart_types = {
            -4100: "xlColumnClustered",
            -4101: "xlColumnStacked",
            -4170: "xlBarClustered",
            -4102: "xlLineStacked",
            73: "xlPie",
            # 더 많은 차트 유형을 필요에 따라 추가
        }
        chart_type_name = chart_types.get(chart_type, f"Unknown Chart Type ({chart_type})")
        
        result += f"\n   - Chart Type: {chart_type_name}"
        result += f"\n   - Has Legend: {chart.HasLegend}"
        result += f"\n   - Has Title: {chart.HasTitle}"
        
        if chart.HasTitle:
            result += f"\n   - Chart Title: {chart.ChartTitle.Text}"
        
        # 차트 데이터 시리즈 정보
        try:
            series_count = chart.SeriesCollection().Count
            result += f"\n   - Series Count: {series_count}"
            
            # 시리즈 이름 예시
            if series_count > 0:
                result += "\n   - Series Names:"
                for i in range(1, min(series_count + 1, 5)):  # 최대 4개 시리즈만 표시
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
    """그룹 내부 Shape 정보 파싱"""
    result = ""
    try:
        shapes_count = group_shape.GroupItems.Count
        result += f"\n   - Group Items Count: {shapes_count}"
        
        # 모든 그룹 아이템을 파싱하면 출력이 너무 길어질 수 있으므로
        # 처음 몇 개 아이템만 간략히 파싱
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
    """이미지 정보 파싱"""
    result = ""
    try:
        result += f"\n   - Picture Type: {picture.Type}"
        
        # 이미지 크기 비율
        if hasattr(picture, 'ScaleHeight') and hasattr(picture, 'ScaleWidth'):
            result += f"\n   - Scale: Width {picture.ScaleWidth}%, Height {picture.ScaleHeight}%"
        
        # 이미지 품질 및 압축 정보
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
    """Placeholder 상세 정보 파싱"""
    result = ""
    try:
        ph_type = placeholder.PlaceholderFormat.Type
        result += f"\n   - Placeholder Type: {get_placeholder_type(ph_type)} (ID: {ph_type})"
        result += f"\n   - Placeholder ID: {placeholder.Id}"
        result += f"\n   - Placeholder Index: {placeholder.PlaceholderFormat.Index}"
        
        # 추가 Placeholder 속성
        if hasattr(placeholder.PlaceholderFormat, 'ContainedType'):
            result += f"\n   - Contained Type: {placeholder.PlaceholderFormat.ContainedType}"
    except Exception as e:
        result += f"\n   - Placeholder Parsing Error: {e}"
    return result

def parse_shape_details(shape):
    """Shape 유형별 세부 정보 파싱"""
    result = ""
    
    # 공통 추가 속성
    try:
        result += f"\n Visibility: {'Visible' if shape.Visible else 'Hidden'}"
        result += f"\n Z-Order: {shape.ZOrderPosition}"
        if hasattr(shape, 'Rotation'):
            result += f"\n Rotation: {shape.Rotation}°"
        result += f"\n ID: {shape.Id}"  # 고유 ID
        
        # 투명도 정보
        if hasattr(shape, 'Fill') and hasattr(shape.Fill, 'Transparency'):
            result += f"\n Fill Transparency: {shape.Fill.Transparency * 100}%"
        
        # 선 정보
        if hasattr(shape, 'Line'):
            line = shape.Line
            if hasattr(line, 'Visible') and line.Visible:
                result += f"\n Line: Width={line.Weight}pt"
                if hasattr(line, 'ForeColor') and hasattr(line.ForeColor, 'RGB'):
                    rgb = line.ForeColor.RGB
                    result += f", Color=RGB({rgb & 0xFF}, {(rgb >> 8) & 0xFF}, {(rgb >> 16) & 0xFF})"
    except Exception as e:
        result += f"\n General Properties Error: {e}"
    
    # Shape 유형별 세부 정보 파싱
    try:
        shape_type = shape.Type
        
        # 텍스트 프레임이 있는 경우
        if hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
            result += parse_text_frame(shape.TextFrame)
        
        # Placeholder인 경우
        if shape_type == 14:  # Placeholder
            result += parse_placeholder_details(shape)
        
        # 그룹인 경우
        elif shape_type == 6:  # Group
            result += parse_group_shapes(shape)
        
        # 테이블인 경우
        elif shape_type == 18:  # Table
            result += parse_table(shape.Table)
        
        # 차트인 경우
        elif shape_type == 3:  # Chart
            result += parse_chart(shape.Chart)
        
        # 이미지인 경우
        elif shape_type in [11, 13]:  # LinkedPicture or Picture
            result += parse_picture(shape)
        
        # SmartArt인 경우
        elif shape_type == 19:  # SmartArt
            result += f"\n   - SmartArt: {shape.SmartArt.AllNodes.Count} nodes"
        
        # OLE 객체인 경우
        elif shape_type in [7, 10]:  # EmbeddedOLEObject or LinkedOLEObject
            result += f"\n   - OLE Class Type: {shape.OLEFormat.ProgID if hasattr(shape, 'OLEFormat') else 'Unknown'}"
        
        # 미디어인 경우
        elif shape_type == 15:  # MediaObject
            result += f"\n   - Media Type: {shape.MediaType if hasattr(shape, 'MediaType') else 'Unknown'}"
    
    except Exception as e:
        result += f"\n Shape Detail Error: {e}"
    
    return result

def parse_slide_notes(slide):
    """슬라이드 노트 파싱"""
    result = "\n\n--- SLIDE NOTES ---"
    try:
        # 노트 페이지 접근
        if hasattr(slide, 'HasNotesPage') and slide.HasNotesPage:
            notes_page = slide.NotesPage
            shapes_count = notes_page.Shapes.Count
            
            result += f"\nNotes Shapes Count: {shapes_count}"
            
            # 노트 페이지의 텍스트 프레임을 찾아 내용 추출
            note_text = ""
            for i in range(1, shapes_count + 1):
                shape = notes_page.Shapes(i)
                if hasattr(shape, 'PlaceholderFormat'):
                    # 노트 텍스트는 보통 PlaceholderType = 2 (Body)
                    if shape.PlaceholderFormat.Type == 2:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            note_text += shape.TextFrame.TextRange.Text
            
            if note_text:
                result += f"\nNotes Content: {note_text}"
            else:
                result += "\nNo text found in notes"
        else:
            result += "\nNo notes page found"
    except Exception as e:
        result += f"\nError parsing notes: {e}"
    
    return result

def parse_slide_properties(slide):
    """슬라이드 속성 파싱"""
    result = "\n\n--- SLIDE PROPERTIES ---"
    try:
        # 슬라이드 기본 정보
        result += f"\nSlide ID: {slide.SlideID}"
        result += f"\nSlide Index: {slide.SlideIndex}"
        result += f"\nSlide Name: {slide.Name}"
        
        # 슬라이드 레이아웃 정보
        if hasattr(slide, 'Layout'):
            result += f"\nSlide Layout Type: {slide.Layout.Type}"
            result += f"\nSlide Layout Name: {slide.Layout.Name}"
        
        # 슬라이드 배경 정보
        if hasattr(slide, 'Background'):
            bg = slide.Background
            if hasattr(bg, 'Fill'):
                fill = bg.Fill
                fill_type = "None"
                if hasattr(fill, 'Type'):
                    fill_types = {1: "Solid", 2: "Pattern", 3: "Gradient", 4: "Texture", 5: "Picture"}
                    fill_type = fill_types.get(fill.Type, f"Unknown ({fill.Type})")
                result += f"\nBackground Fill Type: {fill_type}"
        
        # 슬라이드 전환 효과
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
    """슬라이드 객체 파싱 메인 함수"""
    output = {} # 출력을 저장할 문자열 초기화
    
    try:
        # Connect to running PowerPoint instance
        ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
        # Get active presentation
        presentation = ppt.ActivePresentation
        
        # Check if there is an active presentation
        if not presentation:
            output['status'] = "No active presentation found."
            return output['status']
        
        # 프레젠테이션 정보 추가
        output["Presntation_Name"] = f"{presentation.Name}\n"
        output["Total_Slide_Number"] = f"{presentation.Slides.Count}\n"
        
        # 슬라이드 범위 확인
        if slide_num > presentation.Slides.Count or slide_num < 1:
            output["status"] = f"Invalid slide number. Please provide a number between 1 and {presentation.Slides.Count}."
            return output["status"]
        
        # Access the specified slide
        slide = presentation.Slides(slide_num)
        
        # 슬라이드 속성 파싱
        output["Slide_Properties"] = parse_slide_properties(slide)
        
        # Get the number of shapes in the slide
        shape_count = slide.Shapes.Count
        output["Objects_Overview"] = f"Found {shape_count} objects in slide number {slide_num}."
        
        # Iterate through each shape
        for i in range(1, shape_count + 1):
            shape = slide.Shapes(i)
            output["Objects_Detail"] = []
            output["Objects_Detail"].append(f"Object {i}:")
            output += f"\n Name: {shape.Name}"
            output += f"\n Type: {get_shape_type(shape.Type)}"
            output += f"\n Position: Left={shape.Left}, Top={shape.Top}"
            output += f"\n Size: Width={shape.Width}, Height={shape.Height}"
            
            # Parse details based on shape type
            output += parse_shape_details(shape)
        
        # 슬라이드 노트 파싱
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
    """모든 유형의 도형에서 텍스트를 추출하는 함수"""
    output = ""
    indent = "  " * indent_level
    
    try:
        # 텍스트프레임이 있는 객체에서 텍스트 추출 (TextBox, AutoShape, Placeholder 등)
        if hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
            text_frame = shape.TextFrame
            if hasattr(text_frame, 'HasText') and text_frame.HasText:
                text_range = text_frame.TextRange
                output += f"\n{indent}Text content: {text_range.Text}"
                
                # 텍스트 서식 정보 추출
                try:
                    output += f"\n{indent}Font: {text_range.Font.Name}, Size: {text_range.Font.Size}"
                    output += f"\n{indent}Bold: {text_range.Font.Bold}, Italic: {text_range.Font.Italic}"
                    
                    # 단락 정보 추출
                    if hasattr(text_range, 'ParagraphFormat'):
                        para_format = text_range.ParagraphFormat
                        output += f"\n{indent}Alignment: {get_alignment_type(para_format.Alignment)}"
                        output += f"\n{indent}Line Spacing: {para_format.LineSpacing}"
                except:
                    output += f"\n{indent}Cannot retrieve all text formatting details"
        
        # TextFrame2 지원 (Office 2007 이상)
        elif hasattr(shape, 'HasTextFrame2') and shape.HasTextFrame2:
            text_frame = shape.TextFrame2
            if hasattr(text_frame, 'TextRange') and text_frame.TextRange.Text != "":
                text_range = text_frame.TextRange
                output += f"\n{indent}Text content (TextFrame2): {text_range.Text}"
                
                # 텍스트 서식 정보 추출
                try:
                    output += f"\n{indent}Font: {text_range.Font.Name}, Size: {text_range.Font.Size}"
                    output += f"\n{indent}Bold: {text_range.Font.Bold}, Italic: {text_range.Font.Italic}"
                except:
                    output += f"\n{indent}Cannot retrieve TextFrame2 formatting details"
        
        # 테이블 셀 내의 텍스트 추출
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
                
        # 차트 내의 텍스트 추출
        elif shape.Type == 3:  # Chart
            try:
                chart = shape.Chart
                if chart.HasTitle:
                    output += f"\n{indent}Chart Title: {chart.ChartTitle.Text}"
                    
                # 축 제목 추출
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
                
        # SmartArt 내의 텍스트 추출
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
    # 단락 정렬 값을 텍스트로 변환
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
    output = ""  # 출력을 저장할 문자열 초기화
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
            
            # 그룹 내 개체의 텍스트 추출
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
    output = ""  # 출력을 저장할 문자열 초기화
    indent = "  " * indent_level
    
    # 모든 도형에서 텍스트 추출 시도
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