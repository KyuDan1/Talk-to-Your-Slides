import re
import os
import time  # 시간 관련 기능 추가

class VBAToPythonConverter:
    """PowerPoint VBA 코드를 Python으로 변환하는 클래스"""
    
    def __init__(self):
        self.indent_level = 0
        self.indent_str = "    "  # Python 들여쓰기는 4 공백 사용
        self.python_code = []
        self.vba_section = ""
        self.in_sub = False
        self.current_sub = ""
        self.constant_map = {}
        self.main_function = ""  # 메인 함수 이름을 저장할 변수 추가
        
        # VBA에서 Python으로 변환 규칙 정의
        self.rules = [
            # 기본 임포트 추가
            (r'^Sub\s+(\w+)', self._handle_sub_start),
            (r'^End Sub', self._handle_sub_end),
            (r'\s*Const\s+(\w+)\s+As\s+(\w+)\s+=\s+(.+)', self._handle_const),
            (r'\s*Dim\s+(\w+)\s+As\s+(\w+)', self._handle_dim),
            (r'\s*Set\s+(\w+)\s+=\s+CreateObject\("(.+)"\)', self._handle_createobject),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.(\w+)\.Add', self._handle_add),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.(\w+)\.AddTextbox', self._handle_addtextbox),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.(\w+)\.AddPicture', self._handle_addpicture),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.(\w+)\.AddShape', self._handle_addshape),
            (r'\s*(\w+)\.(\w+)\s+=\s+(.+)', self._handle_property_set),
            (r'\s*(\w+)\.(\w+)\.(\w+)\s+=\s+(.+)', self._handle_nested_property_set),
            (r'\s*(\w+)\.InsertAfter\s+"(.+)"', self._handle_insertafter),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Characters', self._handle_characters),
            (r'\s*(\w+)\.Font\.(\w+)\s+=\s+(.+)', self._handle_font_property),
            (r'\s*(\w+)\.ParagraphFormat\.(\w+)\s+=\s+(.+)', self._handle_paragraph_format),
            (r'\s*(\w+)\.TextFrame\.(\w+)\s+=\s+(.+)', self._handle_textframe_property),
            (r'\s*(\w+)\.InsertAfter\s+vbCrLf', self._handle_insertafter_crlf),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Paragraphs\((\w+)\.Paragraphs\.Count\)', self._handle_last_paragraph),
            (r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Paragraphs\((\d+)\)', self._handle_specific_paragraph),
            (r'\s*MsgBox\s+"(.+)"', self._handle_msgbox),
            # 주석 처리
            (r'\s*\'(.+)', lambda m: f"# {m.group(1)}"),
            # 기타 변환 규칙 추가...
        ]
        
        # VBA 상수를 Python 상수로 매핑
        self.vba_constants = {
            "msoTrue": "-1",
            "msoFalse": "0",
            "CENTER": "2",
            "MIDDLE": "3",
            "TEXT_TO_FIT_SHAPE": "2",
            "SHAPE_TO_FIT_TEXT": "1",
            "NONE": "0",
            "JUSTIFY": "4",
            "msoTextOrientationHorizontal": "1",
            "msoShapeRectangle": "1",
            "ppLayoutBlank": "12",
            "ppLayoutText": "2",
            "ppLayoutTitle": "1",
            "ppLayoutTitleAndContent": "7",
            # PowerPoint TextFrame AutoSize 상수 추가
            "ppAutoSizeNone": "0",
            "ppAutoSizeShapeToFitText": "1",
            "ppAutoSizeTextToFitShape": "2"
        }
    
    def _handle_sub_start(self, match):
        """Sub 선언 처리"""
        sub_name = match.group(1)
        self.in_sub = True
        self.current_sub = sub_name
        self.add_line(f"def {sub_name}(save_path=\"{sub_name}_output.pptx\"):")
        self.indent_level += 1
        self.add_line(f"\"\"\"\n{self.indent_str * self.indent_level}{sub_name} 함수를 통해 PowerPoint 프레젠테이션을 생성합니다.\n{self.indent_str * self.indent_level}\n{self.indent_str * self.indent_level}Args:\n{self.indent_str * self.indent_level}    save_path (str): 저장할 파일 경로\n{self.indent_str * self.indent_level}\"\"\"\n")
        self.add_line("# PowerPoint 애플리케이션 시작")
        self.add_line("print(\"PowerPoint 애플리케이션 시작 중...\")")
        
        # 사용할 상수 추가
        self.add_line("# PowerPoint 상수 정의")
        self.add_line("ppLayoutBlank = 12")
        self.add_line("ppLayoutText = 2")
        self.add_line("ppLayoutTitle = 1")
        self.add_line("ppLayoutTitleAndContent = 7")
        self.add_line("msoTextOrientationHorizontal = 1")
        self.add_line("msoShapeRectangle = 1  # 사각형 도형 상수")
        self.add_line("")
        
        # 텍스트 프레임 설정 함수 추가
        self.add_line("def configure_text_frame(shape):")
        self.indent_level += 1
        self.add_line("\"\"\"텍스트 프레임 설정을 안전하게 수행하는 유틸리티 함수\"\"\"")
        self.add_line("try:")
        self.indent_level += 1
        self.add_line("# 텍스트 프레임 기본 설정")
        self.add_line("shape.TextFrame.WordWrap = -1  # msoTrue")
        self.add_line("shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # ppAlignCenter")
        self.indent_level -= 1
        self.add_line("except Exception as e:")
        self.indent_level += 1
        self.add_line("print(f\"텍스트 프레임 설정 중 오류: {str(e)}\")")
        self.add_line("# 오류 발생 시 대체 방법 시도")
        self.add_line("try:")
        self.indent_level += 1
        self.add_line("# 텍스트 프레임의 텍스트 범위만 설정")
        self.add_line("shape.TextFrame.TextRange.Font.Bold = 0  # msoFalse")
        self.add_line("shape.TextFrame.TextRange.Font.Size = 18")
        self.indent_level -= 1
        self.add_line("except Exception:")
        self.indent_level += 1
        self.add_line("print(\"텍스트 속성 설정 실패\")")
        self.indent_level -= 1
        self.indent_level -= 1
        self.indent_level -= 1
        self.add_line("")
        
        return None
    
    def _handle_sub_end(self, match):
        """Sub 종료 처리"""
        self.in_sub = False
        current_sub = self.current_sub  # 현재 함수 이름 저장
        
        # 파일 저장 로직 추가
        self.add_line("")
        self.add_line("try:")
        self.indent_level += 1
        self.add_line("print(\"프레젠테이션 저장 중...\")")
        self.add_line("")
        self.add_line("# 임시 디렉토리에 저장 경로 설정")
        self.add_line("temp_dir = tempfile.gettempdir()")
        self.add_line(f"temp_filename = f\"{current_sub}_{{int(time.time())}}.pptx\"")
        self.add_line("save_path = os.path.join(temp_dir, temp_filename)")
        self.add_line("")
        self.add_line("# 프레젠테이션 저장")
        self.add_line("ppPres.SaveAs(save_path)")
        self.add_line(f"print(f\"프레젠테이션이 다음 위치에 저장되었습니다: {{save_path}}\")")
        self.add_line("")
        self.add_line("return ppPres, save_path")
        self.indent_level -= 1
        
        self.add_line("except Exception as e:")
        self.indent_level += 1
        self.add_line(f"print(f\"저장 중 오류 발생: {{str(e)}}\")")
        self.add_line("")
        self.add_line("# 대안으로 사용자에게 직접 저장하도록 안내")
        self.add_line("print(\"PowerPoint 애플리케이션이 열려 있습니다. 직접 [파일 > 다른 이름으로 저장]을 선택하여 저장하세요.\")")
        self.add_line("")
        self.add_line("return ppPres, None")
        self.indent_level -= 1
        
        self.indent_level -= 1
        
        # 메인 블록은 별도로 생성
        self.main_function = current_sub  # 메인 블록을 위해 함수 이름 저장
        return None
    
    def _handle_const(self, match):
        """상수 선언 처리"""
        const_name = match.group(1)
        const_value = match.group(3)
        
        # 상수 맵에 저장
        self.constant_map[const_name] = const_value
        
        self.add_line(f"{const_name} = {const_value}")
        return None
    
    def _handle_dim(self, match):
        """변수 선언 처리"""
        var_name = match.group(1)
        var_type = match.group(2)
        if var_type.lower() == "object":
            self.add_line(f"# {var_type} 타입의 {var_name} 선언")
            self.add_line(f"{var_name} = None")
        else:
            self.add_line(f"# {var_type} 타입의 {var_name} 선언")
            self.add_line(f"{var_name} = None")
        return None
    
    def _handle_createobject(self, match):
        """CreateObject 처리"""
        var_name = match.group(1)
        object_type = match.group(2)
        self.add_line(f"{var_name} = win32com.client.Dispatch(\"{object_type}\")")
        return None
    
    def _handle_add(self, match):
        """Add 메서드 처리"""
        var_name = match.group(1)
        parent = match.group(2)
        collection = match.group(3)
        
        # Shapes.Add() 메서드는 파워포인트에서 지원하지 않으므로 분기 처리
        if collection.lower() == "shapes":
            self.add_line(f"# Shapes.Add() 메서드는 파워포인트에서 지원하지 않음")
            self.add_line(f"# 대신 AddShape() 메서드를 사용해야 함")
            self.add_line(f"{var_name} = {parent}.{collection}.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가")
        else:
            self.add_line(f"{var_name} = {parent}.{collection}.Add()")
        return None
    
    def _handle_addtextbox(self, match):
        """AddTextbox 메서드 처리"""
        # 메서드 구현은 convert_line 메서드에서 처리
        return None
    
    def _handle_addpicture(self, match):
        """AddPicture 메서드 처리"""
        # 메서드 구현은 convert_line 메서드에서 처리
        return None
    
    def _handle_addshape(self, match):
        """AddShape 메서드 처리"""
        # 메서드 구현은 convert_line 메서드에서 처리
        return None
    
    def _handle_property_set(self, match):
        """속성 설정 처리"""
        obj_name = match.group(1)
        property_name = match.group(2)
        value = match.group(3)
        
        # 괄호로 묶인 부분 제거 (예: "2 (2)" -> "2")
        value = re.sub(r'(\d+)\s*\(\d+\)', r'\1', value)
        
        # 상수 변환
        for const_name, const_value in self.constant_map.items():
            if const_name in value:
                value = value.replace(const_name, const_value)
        
        # VBA 상수 변환
        for vba_const, py_const in self.vba_constants.items():
            if f"{vba_const}" in value:
                value = value.replace(f"{vba_const}", py_const)
        
        # 슬라이드 크기 설정 특별 처리
        if obj_name.endswith("PageSetup") and (property_name == "SlideWidth" or property_name == "SlideHeight"):
            # VBA는 매우 큰 값으로 슬라이드 크기를 설정하는 경우가 있어, 표준 크기로 변환
            try:
                if int(value) > 10000:  # 매우 큰 값 (VBA 단위)
                    # 표준 와이드스크린 크기로 변환
                    if property_name == "SlideWidth":
                        value = "960"  # 와이드스크린 표준 너비
                    else:
                        value = "540"  # 와이드스크린 표준 높이
            except ValueError:
                # 숫자가 아닌 경우 그대로 사용
                pass
        
        self.add_line(f"{obj_name}.{property_name} = {value}")
        return None
    
    def _handle_nested_property_set(self, match):
        """중첩 속성 설정 처리"""
        obj_name = match.group(1)
        prop1 = match.group(2)
        prop2 = match.group(3)
        value = match.group(4)
        
        # 괄호로 묶인 부분 제거 (예: "2 (2)" -> "2")
        value = re.sub(r'(\d+)\s*\(\d+\)', r'\1', value)
        
        # 상수 변환
        for const_name, const_value in self.constant_map.items():
            if const_name in value:
                value = value.replace(const_name, const_value)
        
        # VBA 상수 변환
        for vba_const, py_const in self.vba_constants.items():
            if f"{vba_const}" in value:
                value = value.replace(f"{vba_const}", py_const)
        
        # TextFrame.AutoSize 속성 처리
        if prop1 == "TextFrame" and prop2 == "AutoSize":
            self.add_line(f"# TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용")
            self.add_line(f"configure_text_frame({obj_name})")
            return None
        
        self.add_line(f"{obj_name}.{prop1}.{prop2} = {value}")
        return None
    
    def _handle_insertafter(self, match):
        """InsertAfter 메서드 처리"""
        obj_name = match.group(1)
        text = match.group(2)
        self.add_line(f"{obj_name}.InsertAfter(\"{text}\")")
        return None
    
    def _handle_insertafter_crlf(self, match):
        """개행 문자 InsertAfter 처리"""
        obj_name = match.group(1)
        self.add_line(f"{obj_name}.InsertAfter(\"\\r\\n\")")
        return None
    
    def _handle_characters(self, match):
        """Characters 메서드 처리"""
        var_name = match.group(1)
        parent = match.group(2)
        self.add_line(f"{var_name} = {parent}.Characters")
        return None
    
    def _handle_font_property(self, match):
        """Font 속성 설정 처리"""
        obj_name = match.group(1)
        property_name = match.group(2)
        value = match.group(3)
        
        # 상수 변환
        for const_name, const_value in self.constant_map.items():
            if const_name in value:
                value = value.replace(const_name, const_value)
        
        # VBA 상수 변환
        for vba_const, py_const in self.vba_constants.items():
            if f"{vba_const}" in value:
                value = value.replace(f"{vba_const}", py_const)
        
        self.add_line(f"try:")
        self.indent_level += 1
        self.add_line(f"{obj_name}.Font.{property_name} = {value}")
        self.indent_level -= 1
        self.add_line(f"except Exception as e:")
        self.indent_level += 1
        self.add_line(f"print(f\"폰트 속성 {property_name} 설정 중 오류: {{str(e)}}\")")
        self.indent_level -= 1
        return None
    
    def _handle_paragraph_format(self, match):
        """ParagraphFormat 속성 설정 처리"""
        obj_name = match.group(1)
        property_name = match.group(2)
        value = match.group(3)
        
        # 상수 변환
        for const_name, const_value in self.constant_map.items():
            if const_name in value:
                value = value.replace(const_name, const_value)
        
        # VBA 상수 변환
        for vba_const, py_const in self.vba_constants.items():
            if f"{vba_const}" in value:
                value = value.replace(f"{vba_const}", py_const)
        
        self.add_line(f"try:")
        self.indent_level += 1
        self.add_line(f"{obj_name}.ParagraphFormat.{property_name} = {value}")
        self.indent_level -= 1
        self.add_line(f"except Exception as e:")
        self.indent_level += 1
        self.add_line(f"print(f\"단락 서식 {property_name} 설정 중 오류: {{str(e)}}\")")
        self.indent_level -= 1
        return None
    
    def _handle_textframe_property(self, match):
        """TextFrame 속성 설정 처리"""
        obj_name = match.group(1)
        property_name = match.group(2)
        value = match.group(3)
        
        # AutoSize 속성은 configure_text_frame 함수를 사용하여 대체
        if property_name == "AutoSize":
            self.add_line(f"# TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용")
            self.add_line(f"configure_text_frame({obj_name})")
            return None
        
        # 괄호로 묶인 부분 제거 (예: "2 (2)" -> "2")
        value = re.sub(r'(\d+)\s*\(\d+\)', r'\1', value)
        
        # 상수 변환
        for const_name, const_value in self.constant_map.items():
            if const_name in value:
                value = value.replace(const_name, const_value)
        
        # VBA 상수 변환
        for vba_const, py_const in self.vba_constants.items():
            if f"{vba_const}" in value:
                value = value.replace(f"{vba_const}", py_const)
        
        # TextRange 설정 관련 특별 처리
        if property_name == "TextRange":
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{obj_name}.TextFrame.TextRange = {value}")
            self.indent_level -= 1
            self.add_line(f"except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"텍스트 범위 설정 중 오류: {{str(e)}}\")")
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{obj_name}.TextFrame.TextRange.Text = {value}")
            self.indent_level -= 1
            self.add_line(f"except Exception as e2:")
            self.indent_level += 1
            self.add_line(f"print(f\"텍스트 설정 중 오류: {{str(e2)}}\")")
            self.indent_level -= 1
            self.indent_level -= 1
        else:
            # 다른 TextFrame 속성에 대한 처리
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{obj_name}.TextFrame.{property_name} = {value}")
            self.indent_level -= 1
            self.add_line(f"except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"텍스트 프레임 {property_name} 설정 중 오류: {{str(e)}}\")")
            self.indent_level -= 1
        
        return None
    
    def _handle_last_paragraph(self, match):
        """마지막 단락 처리"""
        var_name = match.group(1)
        parent = match.group(2)
        self.add_line(f"{var_name} = {parent}.Paragraphs({parent}.Paragraphs.Count)")
        return None
    
    def _handle_specific_paragraph(self, match):
        """특정 단락 처리"""
        var_name = match.group(1)
        parent = match.group(2)
        index = match.group(3)
        self.add_line(f"{var_name} = {parent}.Paragraphs({index})")
        return None
    
    def _handle_msgbox(self, match):
        """MsgBox 처리"""
        msg = match.group(1)
        self.add_line(f"print(\"{msg}\")")
        return None
    
    def add_line(self, line):
        """변환된 Python 코드 라인 추가"""
        if line:
            self.python_code.append(f"{self.indent_str * self.indent_level}{line}")
        else:
            self.python_code.append("")
    
    def convert_line(self, vba_line):
        """단일 VBA 코드 라인을 Python으로 변환"""
        # 슬라이드 추가 패턴 처리
        slides_add_match = re.search(r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Slides\.Add\((\d+),\s*(\d+|\w+)\)', vba_line)
        if slides_add_match:
            var_name = slides_add_match.group(1)
            parent = slides_add_match.group(2)
            index = slides_add_match.group(3)
            layout = slides_add_match.group(4)
            
            # 레이아웃 상수 확인
            if layout in self.constant_map:
                layout = self.constant_map[layout]
            
            # 레이아웃에 대한 변수 이름 사용
            if layout == "1":
                layout = "ppLayoutTitle"
            elif layout == "2":
                layout = "ppLayoutText"
            elif layout == "7":
                layout = "ppLayoutTitleAndContent"
            elif layout == "12":
                layout = "ppLayoutBlank"
                
            self.add_line(f"{var_name} = {parent}.Slides.Add({index}, {layout})")
            return True
        
        # 텍스트 프레임의 AutoSize 속성 직접 설정 패턴 처리
        textframe_autosize_match = re.search(r'\s*(\w+)\.TextFrame\.AutoSize\s*=\s*(.+)', vba_line)
        if textframe_autosize_match:
            obj_name = textframe_autosize_match.group(1)
            self.add_line(f"# TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용")
            self.add_line(f"configure_text_frame({obj_name})")
            return True
        
        # AddTextbox 패턴 처리
        addtextbox_match = re.search(r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Shapes\.AddTextbox\(Orientation:?=(\w+),\s*Left:?=(\d+),\s*Top:?=(\d+),\s*Width:?=(\d+),\s*Height:?=(\d+)\)', vba_line)
        if addtextbox_match:
            var_name = addtextbox_match.group(1)
            parent = addtextbox_match.group(2)
            orientation = addtextbox_match.group(3)
            left = addtextbox_match.group(4)
            top = addtextbox_match.group(5)
            width = addtextbox_match.group(6)
            height = addtextbox_match.group(7)
            
            # VBA 상수 변환
            if orientation == "msoTextOrientationHorizontal":
                orientation = "msoTextOrientationHorizontal"
            else:
                for vba_const, py_const in self.vba_constants.items():
                    if orientation == vba_const:
                        orientation = py_const
            
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{var_name} = {parent}.Shapes.AddTextbox({orientation}, {left}, {top}, {width}, {height})")
            self.add_line(f"# 텍스트박스 생성 후 안정적인 속성 설정을 위한 짧은 지연")
            self.add_line(f"time.sleep(0.1)")
            self.add_line(f"# 텍스트 프레임 안전하게 설정")
            self.add_line(f"configure_text_frame({var_name})")
            self.indent_level -= 1
            
            self.add_line(f"except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"텍스트박스 생성 중 오류: {{str(e)}}\")")
            self.indent_level -= 1
            
            return True
        
        # AddShape 패턴 처리
        addshape_match = re.search(r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Shapes\.AddShape\((\w+),\s*(\d+),\s*(\d+),\s*(\d+),\s*(\d+)\)', vba_line)
        if addshape_match:
            var_name = addshape_match.group(1)
            parent = addshape_match.group(2)
            shape_type = addshape_match.group(3)
            left = addshape_match.group(4)
            top = addshape_match.group(5)
            width = addshape_match.group(6)
            height = addshape_match.group(7)
            
            # VBA 상수 변환
            if shape_type == "msoShapeRectangle":
                shape_type = "msoShapeRectangle"
            else:
                for vba_const, py_const in self.vba_constants.items():
                    if shape_type == vba_const:
                        shape_type = py_const
            
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{var_name} = {parent}.Shapes.AddShape({shape_type}, {left}, {top}, {width}, {height})")
            self.add_line(f"# 도형 생성 후 안정적인 속성 설정을 위한 짧은 지연")
            self.add_line(f"time.sleep(0.1)")
            self.indent_level -= 1
            
            self.add_line(f"except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"도형 생성 중 오류: {{str(e)}}\")")
            self.indent_level -= 1
            
            return True
        
        # AddPicture 패턴 처리
        addpicture_match = re.search(r'\s*Set\s+(\w+)\s+=\s+(\w+)\.Shapes\.AddPicture\(\s*FileName:?="(.+)",\s*LinkToFile:?=(\w+),\s*SaveWithDocument:?=(\w+),\s*Left:?=(\d+),\s*Top:?=(\d+),\s*Width:?=(\d+),\s*Height:?=(\d+)\)', vba_line)
        if addpicture_match:
            var_name = addpicture_match.group(1)
            parent = addpicture_match.group(2)
            filename = addpicture_match.group(3)
            linktofile = addpicture_match.group(4)
            savewithdoc = addpicture_match.group(5)
            left = addpicture_match.group(6)
            top = addpicture_match.group(7)
            width = addpicture_match.group(8)
            height = addpicture_match.group(9)
            
            # VBA 상수 변환
            for vba_const, py_const in self.vba_constants.items():
                if linktofile == vba_const:
                    linktofile = py_const
                if savewithdoc == vba_const:
                    savewithdoc = py_const
            
            self.add_line(f"# 이미지 파일 확인")
            self.add_line(f"if os.path.exists(\"{filename}\"):")
            self.indent_level += 1
            self.add_line(f"try:")
            self.indent_level += 1
            self.add_line(f"{var_name} = {parent}.Shapes.AddPicture(\"{filename}\", {linktofile}, {savewithdoc}, {left}, {top}, {width}, {height})")
            self.indent_level -= 1
            self.add_line(f"except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"이미지 추가 중 오류: {{str(e)}}\")")
            self.indent_level -= 1
            self.indent_level -= 1
            return True
        
        # 일반 규칙 적용
        for pattern, handler in self.rules:
            match = re.match(pattern, vba_line)
            if match:
                result = handler(match)
                if result is not None:
                    self.add_line(result)
                return True
        
        # 처리되지 않은 줄은 주석으로 추가
        if vba_line.strip():
            self.add_line(f"# TODO: {vba_line.strip()}")
        return False
    
    def post_process_code(self):
        """변환된 코드에 대한 후처리 작업"""
        # TextFrame.AutoSize 관련 코드 제거
        processed_code = []
        skip_next = False
        
        for i, line in enumerate(self.python_code):
            if skip_next:
                skip_next = False
                continue
                
            # 직접적인 TextFrame.AutoSize 설정 라인 찾기 및 대체
            if ".TextFrame.AutoSize" in line:
                obj_match = re.search(r'(\w+)\.TextFrame\.AutoSize', line)
                if obj_match:
                    obj_name = obj_match.group(1)
                    processed_code.append(f"{self.indent_str * self.indent_level}# TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용")
                    processed_code.append(f"{self.indent_str * self.indent_level}configure_text_frame({obj_name})")
                    skip_next = True  # 다음 줄 건너뛰기
                    continue
            
            # 괄호로 묶인 숫자 패턴 제거 (예: "2 (2)" -> "2")
            line = re.sub(r'(\d+)\s*\(\s*\d+\s*\)', r'\1', line)
            processed_code.append(line)
        
        self.python_code = processed_code
    
    def convert(self, vba_code):
        """전체 VBA 코드를 Python으로 변환"""
        # 코드 초기화
        self.python_code = []
        self.indent_level = 0
        self.main_function = ""  # 메인 함수 이름 초기화
        
        # VBA 코드에서 숫자(숫자) 패턴 먼저 정리
        vba_code = re.sub(r'(\d+)\s*\(\s*\d+\s*\)', r'\1', vba_code)
        
        # 기본 임포트 추가
        self.add_line("import win32com.client")
        self.add_line("import os")
        self.add_line("import tempfile")
        self.add_line("import time")
        self.add_line("import sys")  # 시스템 모듈 추가
        self.add_line("")
        
        # 각 줄 변환
        for line in vba_code.split('\n'):
            self.convert_line(line)
        
        # 후처리 작업
        self.post_process_code()
        
        # 메인 블록 추가 (함수가 정의된 경우에만)
        if self.main_function:
            self.add_line("if __name__ == \"__main__\":")
            self.indent_level += 1
            self.add_line("try:")
            self.indent_level += 1
            self.add_line("# 실행 경로 설정")
            self.add_line(f"save_path = os.path.join(os.path.expanduser(\"~\"), \"Desktop\", \"{self.main_function}.pptx\")")
            self.add_line("")
            self.add_line(f"# 프레젠테이션 생성")
            self.add_line(f"pres, saved_path = {self.main_function}(save_path)")
            self.add_line("")
            self.add_line(f"if saved_path:")
            self.indent_level += 1
            self.add_line(f"print(f\"프레젠테이션이 다음 위치에 저장되었습니다: {{saved_path}}\")")
            self.indent_level -= 1
            self.indent_level -= 1
            
            self.add_line("except Exception as e:")
            self.indent_level += 1
            self.add_line(f"print(f\"오류 발생: {{str(e)}}\")")
            self.add_line(f"print(\"상세 오류 정보:\")")
            self.add_line(f"import traceback")
            self.add_line(f"traceback.print_exc()")
            self.indent_level -= 1
            self.indent_level -= 1
        
        return '\n'.join(self.python_code)
    
    def convert_file(self, input_file, output_file=None):
        """VBA 파일을 Python 파일로 변환"""
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"입력 파일 {input_file}을 찾을 수 없습니다.")
        
        with open(input_file, 'r', encoding='utf-8') as f:
            vba_code = f.read()
        
        python_code = self.convert(vba_code)
        
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(python_code)
            print(f"{output_file}에 변환된 코드가 저장되었습니다.")
        
        return python_code


# 사용 예제
if __name__ == "__main__":
    converter = VBAToPythonConverter()
    
    # 예제: 파일에서 변환
    input_file = "TeXBLEU_ppt_vba.bas"  # VBA 코드 파일 경로
    output_file = "python_code.py"  # 변환된 Python 코드 저장 경로
    
    try:
        # 커맨드 라인 인자로 받을 수도 있음
        converter.convert_file(input_file, output_file)
    except FileNotFoundError as e:
        print(f"오류: {e}")
        print("예제 코드 변환 실행...")
        
        # 예제 코드로 변환 데모
        example_vba = """Sub CreatePresentationFromAnalysis()
    ' PowerPoint 애플리케이션 객체 생성
    Dim ppApp As Object
    Dim ppPres As Object
    Dim ppSlide As Object
    
    ' PowerPoint 상수 선언
    Const msoTrue As Long = -1
    Const msoFalse As Long = 0
    
    ' PowerPoint 시작
    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = True 
    
    ' 새 프레젠테이션 생성
    Set ppPres = ppApp.Presentations.Add
    
    ' 슬라이드 크기 설정
    ppPres.PageSetup.SlideWidth = 12192000
    ppPres.PageSetup.SlideHeight = 6858000

    ' 슬라이드 추가
    Set ppSlide = ppPres.Slides.Add(1, 7)
    
    ' 텍스트 상자 추가
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=100, Top:=100, Width:=400, Height:=100)
    shp.TextFrame.TextRange.Text = "Hello, World!"
    shp.TextFrame.AutoSize = 2
    
    MsgBox "프레젠테이션이 성공적으로 생성되었습니다."
End Sub"""
        
        converted_code = converter.convert(example_vba)
        print("\n변환 결과:")
        print(converted_code)
        
        # 결과 파일로 저장
        with open("example_converted.py", "w", encoding="utf-8") as f:
            f.write(converted_code)
        print("example_converted.py 파일로 변환 결과가 저장되었습니다.")