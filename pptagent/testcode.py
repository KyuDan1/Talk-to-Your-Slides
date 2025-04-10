import win32com.client
import pywintypes

def parse_active_slide_objects():
    try:
        # 실행 중인 PowerPoint 인스턴스에 연결
        ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
        # 활성화된 프레젠테이션 가져오기
        presentation = ppt.ActivePresentation
        
        # 활성화된 프레젠테이션이 있는지 확인
        if not presentation:
            print("활성화된 프레젠테이션이 없습니다.")
            return
        
        # 첫 번째 슬라이드 접근
        slide = presentation.Slides(1)
        
        # 슬라이드에 있는 도형 개수 가져오기
        shape_count = slide.Shapes.Count
        print(f"첫 번째 슬라이드에서 {shape_count}개의 객체를 발견했습니다.")
        
        # 각 도형을 순회
        for i in range(1, shape_count + 1):
            shape = slide.Shapes(i)
            print(f"\n객체 {i}:")
            print(f"  이름: {shape.Name}")
            print(f"  유형: {get_shape_type(shape.Type)}")
            print(f"  위치: 왼쪽={shape.Left}, 위쪽={shape.Top}")
            print(f"  크기: 너비={shape.Width}, 높이={shape.Height}")
            
            # 도형 유형에 따라 세부 정보 파싱
            parse_shape_details(shape)
                
        print("\n파싱 완료.")
        
    except pywintypes.com_error as e:
        print(f"COM 오류: {e}")
    except Exception as e:
        print(f"오류: {e}")

def get_shape_type(type_val):
    # 도형 유형 값을 읽기 쉬운 이름으로 매핑
    # 공식 문서: https://learn.microsoft.com/en-us/office/vba/api/office.msoshapetype
    shape_types = {
        1: "자동 도형",
        2: "설명선",
        3: "차트",
        4: "메모",
        5: "자유형",
        6: "그룹",
        7: "임베디드 OLE 객체",
        8: "폼 컨트롤",
        9: "선",
        10: "연결된 OLE 객체",
        11: "연결된 그림",
        12: "OLE 컨트롤 객체",
        13: "그림",
        14: "자리 표시자",
        15: "텍스트 효과",
        16: "미디어",
        17: "텍스트 상자",
        18: "스크립트 앵커",
        19: "표",
        20: "캔버스",
        21: "다이어그램",
        22: "잉크",
        23: "잉크 메모",
        24: "스마트 아트",
        25: "웹 비디오",
        26: "콘텐츠 앱"
    }
    return shape_types.get(type_val, f"알 수 없는 유형 ({type_val})")

def parse_shape_details(shape):
    # 도형 유형에 따라 특정 세부 정보 추출
    try:
        if shape.Type == 17:  # 텍스트 상자
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                if text_frame.HasText:
                    text_range = text_frame.TextRange
                    print(f"  텍스트 내용: {text_range.Text}")
                    
                    # 텍스트 서식 세부 정보 가져오기
                    try:
                        print(f"  글꼴: {text_range.Font.Name}, 크기: {text_range.Font.Size}")
                        print(f"  굵게: {text_range.Font.Bold}, 기울임꼴: {text_range.Font.Italic}")
                    except:
                        print("  모든 텍스트 서식 세부 정보를 가져올 수 없습니다")
        
        elif shape.Type == 13:  # 그림
            print(f"  그림: {shape.Name}")
            try:
                print(f"  대체 텍스트: {shape.AlternativeText}")
            except:
                pass
                
        elif shape.Type == 3:  # 차트
            print(f"  차트: {shape.Name}")
            try:
                chart = shape.Chart
                print(f"  차트 유형: {chart.ChartType}")
                print(f"  제목 포함 여부: {chart.HasTitle}")
                if chart.HasTitle:
                    print(f"  제목: {chart.ChartTitle.Text}")
            except:
                print("  모든 차트 세부 정보를 가져올 수 없습니다")
                
        elif shape.Type == 19:  # 표
            print(f"  표: {shape.Name}")
            try:
                table = shape.Table
                print(f"  행: {table.Rows.Count}, 열: {table.Columns.Count}")
            except:
                print("  모든 표 세부 정보를 가져올 수 없습니다")
                
    except Exception as e:
        print(f"  도형 세부 정보 파싱 오류: {e}")

if __name__ == "__main__":
    parse_active_slide_objects()