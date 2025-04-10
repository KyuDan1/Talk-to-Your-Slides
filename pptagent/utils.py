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
    try:
        if shape.Type == 6:  # Group
            output += f"\n{indent}Group object found: {shape.Name}"
            output += parse_group_shape(shape, indent_level)
            
        elif shape.Type == 17:  # Text Box
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    text_range = text_frame.TextRange
                    output += f"\n{indent}Text content: {text_range.Text}"
                    
                    # Get text formatting details
                    try:
                        output += f"\n{indent}Font: {text_range.Font.Name}, Size: {text_range.Font.Size}"
                        output += f"\n{indent}Bold: {text_range.Font.Bold}, Italic: {text_range.Font.Italic}"
                    except:
                        output += f"\n{indent}Cannot retrieve all text formatting details"
        
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
                if chart.HasTitle:
                    output += f"\n{indent}Title: {chart.ChartTitle.Text}"
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