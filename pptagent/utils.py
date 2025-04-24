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
    """텍스트 프레임 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        has_text = getattr(text_frame, "HasText", False)
        result["Has Text"] = bool(has_text)
        
        if has_text:
            tr = text_frame.TextRange
            # 기본 텍스트 정보
            result["Text"] = getattr(tr, "Text", "")
            result["Paragraphs Count"] = getattr(tr.Paragraphs(), "Count", 0)
            result["Text Alignment"] = getattr(tr.ParagraphFormat, "Alignment", None)
            
            # 글꼴 정보
            font = tr.Font
            font_info = {
                "Name": getattr(font, "Name", None),
                "Size": getattr(font, "Size", None),
            }
            # RGB 색 추출
            rgb = getattr(font.Color, "RGB", None)
            if rgb is not None:
                font_info["Color"] = {
                    "R": rgb & 0xFF,
                    "G": (rgb >> 8) & 0xFF,
                    "B": (rgb >> 16) & 0xFF
                }
            result["Font"] = font_info

            # 하이퍼링크 검사
            try:
                action = tr.ActionSettings(1)
                hl = getattr(action, "Hyperlink", None)
                addr = getattr(hl, "Address", None) if hl else None
                result["Hyperlink"] = addr or None
            except Exception:
                result["Hyperlink"] = None

    except Exception as e:
        result["Text Frame Error"] = str(e)

    return result


def parse_table(table):
    """테이블 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        rows = getattr(table.Rows, "Count", 0)
        cols = getattr(table.Columns, "Count", 0)
        result["Dimensions"] = {"Rows": rows, "Columns": cols}
        result["FirstRow"]   = getattr(table, "FirstRow", None)
        result["LastRow"]    = getattr(table, "LastRow", None)
        result["FirstCol"]   = getattr(table, "FirstCol", None)
        result["LastCol"]    = getattr(table, "LastCol", None)

        # 샘플 셀 내용
        samples = {}
        max_r = min(3, rows)
        max_c = min(3, cols)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                key = f"Cell({r},{c})"
                try:
                    txt = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                    samples[key] = txt[:30] + ("..." if len(txt) > 30 else "")
                except Exception:
                    samples[key] = None
        result["Sample Cells"] = samples

    except Exception as e:
        result["Table Parsing Error"] = str(e)
    return result


def parse_chart(chart):
    """차트 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        ct = getattr(chart, "ChartType", None)
        chart_types = {
            -4100: "xlColumnClustered", -4101: "xlColumnStacked",
            -4170: "xlBarClustered",    -4102: "xlLineStacked",
             73:    "xlPie"
        }
        result["Chart Type"] = chart_types.get(ct, f"Unknown ({ct})")
        result["Has Legend"] = bool(getattr(chart, "HasLegend", False))
        result["Has Title"]  = bool(getattr(chart, "HasTitle", False))
        if result["Has Title"]:
            result["Title Text"] = getattr(chart.ChartTitle, "Text", None)

        # 시리즈 정보
        series_info = {}
        try:
            sc = chart.SeriesCollection()
            count = getattr(sc, "Count", 0)
            series_info["Count"] = count
            for i in range(1, min(count, 4) + 1):
                try:
                    series_info[f"Series {i} Name"] = sc.Item(i).Name
                except Exception:
                    series_info[f"Series {i} Name"] = None
        except Exception as se:
            series_info["Error"] = str(se)
        result["Series"] = series_info

    except Exception as e:
        result["Chart Parsing Error"] = str(e)
    return result


def parse_group_shapes(group_shape):
    """그룹 내부 Shape 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        count = getattr(group_shape.GroupItems, "Count", 0)
        result["Group Items Count"] = count
        items = {}
        for i in range(1, min(count, 3) + 1):
            try:
                sub = group_shape.GroupItems.Item(i)
                items[f"Item {i}"] = {
                    "Name": getattr(sub, "Name", None),
                    "Type": getattr(sub, "Type", None)
                }
            except Exception:
                items[f"Item {i}"] = None
        result["Items"] = items

    except Exception as e:
        result["Group Parsing Error"] = str(e)
    return result


def parse_picture(picture):
    """이미지 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        result["Type"] = getattr(picture, "Type", None)
        result["Scale"] = {
            "Width %": getattr(picture, "ScaleWidth", None),
            "Height %": getattr(picture, "ScaleHeight", None)
        }
        pf = getattr(picture, "PictureFormat", None)
        if pf:
            pic_fmt = {}
            for attr in ("Brightness", "Contrast"):
                if hasattr(pf, attr):
                    pic_fmt[attr] = getattr(pf, attr)
            crop = getattr(pf, "Crop", None)
            if crop:
                pic_fmt["Crop"] = {
                    "Left": getattr(crop, "ShapeLeft", None),
                    "Top": getattr(crop, "ShapeTop", None),
                    "Width": getattr(crop, "ShapeWidth", None),
                    "Height": getattr(crop, "ShapeHeight", None)
                }
            result["PictureFormat"] = pic_fmt

    except Exception as e:
        result["Picture Parsing Error"] = str(e)
    return result


def parse_placeholder_details(placeholder):
    """Placeholder 상세 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        pf = placeholder.PlaceholderFormat
        ptype = getattr(pf, "Type", None)
        result["Placeholder Type"]  = ptype
        result["Placeholder Type Name"] = get_placeholder_type(ptype)
        result["Placeholder ID"]    = getattr(placeholder, "Id", None)
        result["Placeholder Index"] = getattr(pf, "Index", None)
        if hasattr(pf, "ContainedType"):
            result["Contained Type"] = getattr(pf, "ContainedType", None)

    except Exception as e:
        result["Placeholder Parsing Error"] = str(e)
    return result

def parse_shape_details(shape):
    """Shape 유형별 세부 정보 파싱 (결과를 dict로 반환)"""
    result = {}

    # 공통 속성
    try:
        result["Visibility"] = "Visible" if getattr(shape, "Visible", False) else "Hidden"
        result["Z-Order"]    = getattr(shape, "ZOrderPosition", None)
        if hasattr(shape, "Rotation"):
            result["Rotation (°)"] = getattr(shape, "Rotation", None)
        result["ID"] = getattr(shape, "Id", None)

        # 투명도
        fill = getattr(shape, "Fill", None)
        if fill and hasattr(fill, "Transparency"):
            result["Fill Transparency (%)"] = fill.Transparency * 100

        # 선 정보
        line = getattr(shape, "Line", None)
        if line and getattr(line, "Visible", False):
            line_info = {
                "Width (pt)": getattr(line, "Weight", None)
            }
            fore = getattr(line, "ForeColor", None)
            if fore and hasattr(fore, "RGB"):
                rgb = fore.RGB
                line_info["Color"] = {
                    "R": rgb & 0xFF,
                    "G": (rgb >> 8) & 0xFF,
                    "B": (rgb >> 16) & 0xFF
                }
            result["Line"] = line_info

    except Exception as e:
        result["General Properties Error"] = str(e)

    # 타입별 세부 정보
    try:
        t = getattr(shape, "Type", None)

        # 텍스트 프레임
        if getattr(shape, "HasTextFrame", False):
            result["TextFrame"] = parse_text_frame(shape.TextFrame)

        # Placeholder
        if t == 14:
            result["Placeholder"] = parse_placeholder_details(shape)

        # Group
        elif t == 6:
            result["GroupShapes"] = parse_group_shapes(shape)

        # Table
        elif t == 18:
            result["Table"] = parse_table(shape.Table)

        # Chart
        elif t == 3:
            result["Chart"] = parse_chart(shape.Chart)

        # Picture
        elif t in (11, 13):
            result["Picture"] = parse_picture(shape)

        # SmartArt
        elif t == 19:
            result["SmartArt Nodes"] = getattr(shape.SmartArt.AllNodes, "Count", None)

        # OLE Object
        elif t in (7, 10):
            prog = getattr(shape.OLEFormat, "ProgID", None) if hasattr(shape, "OLEFormat") else None
            result["OLE Class Type"] = prog or "Unknown"

        # Media
        elif t == 15:
            result["Media Type"] = getattr(shape, "MediaType", "Unknown")

    except Exception as e:
        result["Shape Detail Error"] = str(e)

    return result


def parse_slide_notes(slide):
    """슬라이드 노트 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        # 노트 페이지 유무
        has_notes = getattr(slide, "HasNotesPage", False)
        result["Has Notes Page"] = bool(has_notes)

        if has_notes:
            notes_page = slide.NotesPage
            shapes = notes_page.Shapes
            count = getattr(shapes, "Count", 0)
            result["Notes Shapes Count"] = count

            # 노트 텍스트 수집
            texts = []
            for i in range(1, count + 1):
                shape = shapes(i)
                ph = getattr(shape, "PlaceholderFormat", None)
                if ph and getattr(ph, "Type", None) == 2:
                    tf = getattr(shape, "TextFrame", None)
                    if tf and getattr(tf, "HasText", False):
                        texts.append(shape.TextFrame.TextRange.Text)

            # 내용 유무에 따라 설정
            if texts:
                result["Notes Content"] = "".join(texts)
            else:
                result["Notes Content"] = None
        else:
            result["Notes Content"] = None

    except Exception as e:
        result["Error parsing notes"] = str(e)

    return result


def parse_slide_properties(slide):
    """슬라이드 속성 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        # Layout 은 COM 상에서 단순 enum(int) 이므로
        # .Type/.Name 을 호출하면 int 에서 에러가 남.
        # 대신 코드값만 저장하고, CustomLayout 객체를 쓰세요.
        layout_code = getattr(slide, "Layout", None)
        if layout_code is not None:
            result["Slide Layout Code"] = layout_code

        # CustomLayout 은 객체이므로 이름/인덱스 등을 가져올 수 있음
        custom = getattr(slide, "CustomLayout", None)
        if custom is not None:
            result["CustomLayout Name"]  = getattr(custom, "Name", None)
            result["CustomLayout Index"] = getattr(custom, "Index", None)

        # 배경 채우기 정보
        bg = getattr(slide, "Background", None)
        if bg is not None:
            fill = getattr(bg, "Fill", None)
            if fill is not None:
                # fill.Type 은 안전하게 getattr 으로
                t = getattr(fill, "Type", None)
                fill_types = {1: "Solid", 2: "Pattern", 3: "Gradient", 4: "Texture", 5: "Picture"}
                result["Background Fill Type"] = fill_types.get(t, f"Unknown ({t})")

        # 전환 효과
        trans = getattr(slide, "SlideShowTransition", None)
        if trans is not None:
            result["Transition Effect"]   = getattr(trans, "EntryEffect", "None")
            result["Advance Time (s)"]    = getattr(trans, "AdvanceTime", "Manual")
            result["Advance On Click"]    = bool(getattr(trans, "AdvanceOnClick", False))
            result["Advance On Time"]     = bool(getattr(trans, "AdvanceOnTime", False))

    except Exception as e:
        result["error"] = str(e)

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
        output["Presentation_Name"] = f"{presentation.Name}"
        output["Total_Slide_Number"] = f"{presentation.Slides.Count}"
        
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
        output["Objects_Detail"] = []

        # Iterate through each shape
        for i in range(1, shape_count + 1):
            shape = slide.Shapes(i)
            shape_info = {
                "Object_number": i,
                "Name": shape.Name,
                "Type": get_shape_type(shape.Type),
                "Position_Left": shape.Left,
                "Position_Top": shape.Top,
                "Size_Width": shape.Width,
                "Size_Height": shape.Height,
                "More_detail": parse_shape_details(shape),
                
            }
            output["Objects_Detail"].append(shape_info)
        
        # 슬라이드 노트 파싱
        output["Slide_Notes"] = parse_slide_notes(slide)
        
    except pywintypes.com_error as e:
        output["Error"] = f"COM error: {e}"
    # except Exception as e:
    #     output["Error"] = f"Error: {e}"
    
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
    """
    모든 유형의 도형에서 텍스트를 추출 (결과를 dict로 반환)
    """
    result = {}
    # try:
        # 1) TextFrame 지원 객체
    if getattr(shape, "HasTextFrame", False) and shape.TextFrame.HasText:
        tr = shape.TextFrame.TextRange
        font = tr.Font
        result["TextFrame"] = {
            "Text": getattr(tr, "Text", ""),
            "Font": {
                "Name": getattr(font, "Name", None),
                "Size": getattr(font, "Size", None),
                "Bold": getattr(font, "Bold", None),
                "Italic": getattr(font, "Italic", None),
            },
            "Paragraphs": {
                "Count": getattr(tr.Paragraphs(), "Count", 0),
                "Alignment": get_alignment_type(getattr(tr.ParagraphFormat, "Alignment", None)),
                "LineSpacing": getattr(tr.ParagraphFormat, "LineSpacing", None),
            },
            "Hyperlink": None
        }
        # Hyperlink 있으면 추가
        try:
            hl = tr.ActionSettings(1).Hyperlink
            addr = getattr(hl, "Address", None)
            if addr:
                result["TextFrame"]["Hyperlink"] = addr
        except:
            pass

    # 2) TextFrame2 (Office2007+) 지원 객체
    elif getattr(shape, "HasTextFrame2", False) and shape.TextFrame2.TextRange.Text:
        tr2 = shape.TextFrame2.TextRange
        font2 = tr2.Font
        result["TextFrame2"] = {
            "Text": getattr(tr2, "Text", ""),
            "Font": {
                "Name": getattr(font2, "Name", None),
                "Size": getattr(font2, "Size", None),
                "Bold": getattr(font2, "Bold", None),
                "Italic": getattr(font2, "Italic", None),
            }
        }

    # 3) Table 내 텍스트
    elif getattr(shape, "Type", None) == 19 and hasattr(shape, "Table"):
        tbl = shape.Table
        rows, cols = getattr(tbl.Rows, "Count", 0), getattr(tbl.Columns, "Count", 0)
        cells = {}
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                key = f"Cell({r},{c})"
                try:
                    txt = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
                except:
                    txt = None
                cells[key] = txt
        result["TableText"] = {"Rows": rows, "Columns": cols, "Cells": cells}

    # 4) Chart 내 텍스트
    elif getattr(shape, "Type", None) == 3 and hasattr(shape, "Chart"):
        chart = shape.Chart
        chart_info = {
            "Title": getattr(chart.ChartTitle, "Text", None) if getattr(chart, "HasTitle", False) else None,
            "Axes": {}
        }
        if hasattr(chart, "Axes"):
            for grp in (1, 2, 3):
                for typ in (1, 2):
                    try:
                        ax = chart.Axes(grp, typ)
                        if getattr(ax, "HasTitle", False):
                            chart_info["Axes"][f"{grp},{typ}"] = ax.AxisTitle.Text
                    except:
                        pass
        result["ChartText"] = chart_info

    # 5) SmartArt 내 텍스트
    elif getattr(shape, "Type", None) == 24 and hasattr(shape, "SmartArt"):
        nodes = getattr(shape.SmartArt, "AllNodes", None)
        smart = {}
        if nodes:
            for i in range(1, getattr(nodes, "Count", 0) + 1):
                try:
                    txt = nodes.Item(i).TextFrame2.TextRange.Text
                except:
                    txt = None
                smart[f"Node {i}"] = txt
        result["SmartArtText"] = smart

    else:
        result["Text"] = None

    # except Exception as e:
    #     result["Error"] = str(e)

    return result



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
    """
    Shape 유형별 세부 정보 파싱 (결과를 dict로 반환)
    """
    result = {}
    
    # 1) 먼저 텍스트 관련 정보를 dict로 가져와 병합
    text_info = extract_text_from_shape(shape, indent_level)
    if isinstance(text_info, dict):
        result.update(text_info)
    
    # 2) Shape 유형별 추가 정보
    try:
        t = getattr(shape, "Type", None)
        
        # 그룹
        if t == 6:
            grp = {"Group": {
                "Name": getattr(shape, "Name", None),
                "Items": parse_group_shapes(shape)  # 이 함수도 dict 반환 가정
            }}
            result.update(grp)
        
        # 그림(Picture)
        elif t == 13:
            pic_info = {"Picture": {
                "Name": getattr(shape, "Name", None),
                "AlternativeText": getattr(shape, "AlternativeText", None)
            }}
            result.update(pic_info)
        
        # 차트(Chart)
        elif t == 3:
            chart = getattr(shape, "Chart", None)
            chart_info = {"Chart": {
                "Name": getattr(shape, "Name", None),
                "ChartType": getattr(chart, "ChartType", None) if chart else None,
                "HasTitle": getattr(chart, "HasTitle", None) if chart else None
            }}
            result.update(chart_info)
        
        # 테이블(Table)
        elif t == 19:
            table = getattr(shape, "Table", None)
            table_info = {"Table": {
                "Name": getattr(shape, "Name", None),
                "Rows": getattr(table.Rows, "Count", None) if table else None,
                "Columns": getattr(table.Columns, "Count", None) if table else None
            }}
            result.update(table_info)
    
    except Exception as e:
        result["Shape Detail Error"] = str(e)
    
    return result


# output = parse_active_slide_objects()
# print(output)