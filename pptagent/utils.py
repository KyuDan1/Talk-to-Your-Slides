import win32com.client
import pywintypes
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
    
def parse_active_slide_objects(slide_num:int=1):
    output = ""  # 출력을 저장할 문자열 초기화
    try:
        # Connect to running PowerPoint instance
        ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
        # Get active presentation
        presentation = ppt.ActivePresentation
        
        # Check if there is an active presentation
        if not presentation:
            output += "No active presentation found."
            return output
        
        # Access the first slide
        slide = presentation.Slides(slide_num)
        
        # Get the number of shapes in the slide
        shape_count = slide.Shapes.Count
        output += f"Found {shape_count} objects in the slide number {slide_num}."
        
        # Iterate through each shape
        for i in range(1, shape_count + 1):
            shape = slide.Shapes(i)
            output += f"\nObject {i}:"
            output += f"\n  Name: {shape.Name}"
            output += f"\n  Type: {get_shape_type(shape.Type)}"
            output += f"\n  Position: Left={shape.Left}, Top={shape.Top}"
            output += f"\n  Size: Width={shape.Width}, Height={shape.Height}"
            
            # Parse details based on shape type
            output += parse_shape_details(shape)
                
        output += "\nParsing complete."
        
    except pywintypes.com_error as e:
        output += f"COM error: {e}"
    except Exception as e:
        output += f"Error: {e}"
    
    return output

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