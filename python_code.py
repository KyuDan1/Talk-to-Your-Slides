import win32com.client
import os
import tempfile
import time
import sys

def CreatePresentationFromAnalysis(save_path="CreatePresentationFromAnalysis_output.pptx"):
    """
    CreatePresentationFromAnalysis 함수를 통해 PowerPoint 프레젠테이션을 생성합니다.
    
    Args:
        save_path (str): 저장할 파일 경로
    """

    # PowerPoint 애플리케이션 시작
    print("PowerPoint 애플리케이션 시작 중...")
    # PowerPoint 상수 정의
    ppLayoutBlank = 12
    ppLayoutText = 2
    ppLayoutTitle = 1
    ppLayoutTitleAndContent = 7
    msoTextOrientationHorizontal = 1
    msoShapeRectangle = 1  # 사각형 도형 상수

    def configure_text_frame(shape):
        """텍스트 프레임 설정을 안전하게 수행하는 유틸리티 함수"""
        try:
            # 텍스트 프레임 기본 설정
            shape.TextFrame.WordWrap = -1  # msoTrue
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # ppAlignCenter
        except Exception as e:
            print(f"텍스트 프레임 설정 중 오류: {str(e)}")
            # 오류 발생 시 대체 방법 시도
            try:
                # 텍스트 프레임의 텍스트 범위만 설정
                shape.TextFrame.TextRange.Font.Bold = 0  # msoFalse
                shape.TextFrame.TextRange.Font.Size = 18
            except Exception:
                print("텍스트 속성 설정 실패")

    #  PowerPoint 애플리케이션 객체 생성
    # Object 타입의 ppApp 선언
    ppApp = None
    # Object 타입의 ppPres 선언
    ppPres = None
    # Object 타입의 ppSlide 선언
    ppSlide = None
    # Object 타입의 shp 선언
    shp = None
    # Object 타입의 textRange 선언
    textRange = None
    # Object 타입의 paraRange 선언
    paraRange = None
    # Object 타입의 runRange 선언
    runRange = None
    # Object 타입의 tableCell 선언
    tableCell = None
    #  PowerPoint 상수 선언
    msoTrue = -1
    msoFalse = 0
    #  PowerPoint 시작
    ppApp = win32com.client.Dispatch("PowerPoint.Application")
    ppApp.Visible = True
    #  새 프레젠테이션 생성
    ppPres = ppApp.Presentations.Add()
    #  슬라이드 크기 설정
    ppPres.PageSetup.SlideWidth = 12192000
    ppPres.PageSetup.SlideHeight = 6858000
    #  ----- 슬라이드 1 추가 (제목 슬라이드) -----
    ppSlide = ppPres.Slides.Add(1, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=1160584, Top:=884855, _
    # TODO: Width:=9870831, Height:=2387600)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(": Automatic Metric for Evaluate LaTeX Format")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 부제목 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=1956079, Top:=3364568, _
    # TODO: Width:=8279842, Height:=1655762)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Kyudan")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Jung")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Nam-Joon Kim")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2∗")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Hyun Gon Ryu")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("3")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Sieun")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Hyeon")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Seung-")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("jun")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Lee")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Hyuk-jae Lee")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Chung-Ang University, Seoul, Korea")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Seoul National ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Univeristy")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(", Seoul, Korea")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("NVIDIA")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=0, Top:=5270359, _
    # TODO: Width:=1969477, Height:=1587641)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=203977, Top:=260510, _
    # TODO: Width:=1561521, Height:=890067)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=10578549, Top:=170309, _
    # TODO: Width:=1409474, Height:=1042255)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=9334500, Top:=5649971, _
    # TODO: Width:=2857500, Height:=1208029)
    #  텍스트 도형 추가: TextBox 5
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=4655828, Top:=5842391, _
    # TODO: Width:=2893741, Height:=369332)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = 0
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = 2
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("wjdrbeks1021@gmail.com")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  ----- 슬라이드 2 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(2, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=1325563)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Overview")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1527349, _
    # TODO: Width:=10515600, Height:=4965526)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Problem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("1.1 Metric for LaTeX format")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("1.2 Limitations of existing metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("2.1 Tokenizer trained by ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("arXiv")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" papers")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("2.2 Fine-tuned embedding model")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("2.3 Computing ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Experiments")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("3.1 Dataset and setup")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("3.2 Human evaluation")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter("3.3 Result")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Conclusion")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 3
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  ----- 슬라이드 3 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(3, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.Problem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838199, Top:=1825625, _
    # TODO: Width:=10930247, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("LaTeX format")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" : Useful format in Mathematics, Computer Science, etc.")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Example: In converting spoken mathematical expressions into LaTeX")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Command-based language such as C or SQL")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Suitable evaluation metric is required")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.1 Metric for LaTeX Format ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  ----- 슬라이드 4 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(4, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.Promblem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("BLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.2 Limitations of existing metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2381459, Top:=2744905, _
    # TODO: Width:=7429081, Height:=2512778)
    #  ----- 슬라이드 5 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(5, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.Promblem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("sacreBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.2 Limitations of existing metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2424350, Top:=2823116, _
    # TODO: Width:=7343299, Height:=2651408)
    #  ----- 슬라이드 6 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(6, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.Promblem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Character Error Rate (CER)")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.2 Limitations of existing metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2596585, Top:=2897668, _
    # TODO: Width:=6998829, Height:=2340707)
    #  ----- 슬라이드 7 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(7, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.Promblem Define")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Word Error Rate (WER)")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1.2 Limitations of existing metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2338720, Top:=2908720, _
    # TODO: Width:=7514560, Height:=2542054)
    #  ----- 슬라이드 8 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(8, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Requiring ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("LaTeX-specified tokenizer")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Collecting papers in ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("arXiv")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" .")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("tex")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" files")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Byte-pair encoding (BPE) based tokenizer")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Unlike other tokenizers, the tokenizer is designed to ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("capture LaTeX grammar elements")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" and tokenize them in a way that ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("reflects LaTeX’s unique structure")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.1 Tokenizer trained by arXiv papers")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  ----- 슬라이드 9 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(9, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=371809, Top:=4283565, _
    # TODO: Width:=11634143, Height:=1893397)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = 4
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Fig. 4. The results of tokenizing the quadratic formula in LaTeX format using various models. ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("The ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("green boxes indicate sections where LaTeX commands were successfully tokenized as complete chunks")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(", while the ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("red boxes represent sections where the LaTeX commands were fragmented, losing their meaning during tokenization")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(255, 0, 0)
    paraRange.InsertAfter(". It is evident that the tokenizer we developed using the ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("arXiv")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" paper corpus performs the best in accurately tokenizing LaTeX commands.")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.1 Tokenizer trained by arXiv papers")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=647972, Top:=1908518, _
    # TODO: Width:=10981990, Height:=2198512)
    #  ----- 슬라이드 10 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(10, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Based on new tokenizer, we made ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("new embedding model")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Fine-tuning a publicly available pretrained ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("GPT-2 embedding model")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.2 Fine-tuned embedding model")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  ----- 슬라이드 11 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(11, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.3 Computing TeXBLEU argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2461240, Top:=2358845, _
    # TODO: Width:=7520960, Height:=927651)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2960064, Top:=5504189, _
    # TODO: Width:=6920207, Height:=633833)
    #  ----- 슬라이드 12 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(12, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.3 Computing TeXBLEU argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2495537, Top:=1504075, _
    # TODO: Width:=6719715, Height:=5345045)
    #  ----- 슬라이드 13 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(13, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.3 Computing TeXBLEU argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2730326, Top:=4219532, _
    # TODO: Width:=6493842, Height:=1061213)
    #  ----- 슬라이드 14 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(14, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825624, _
    # TODO: Width:=10515600, Height:=4329699)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Finally, ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" is finally equal to the following expression")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.3 Computing TeXBLEU argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2512660, Top:=3230252, _
    # TODO: Width:=7166679, Height:=1265938)
    #  ----- 슬라이드 15 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(15, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.Method")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2.3 Computing TeXBLEU argorithm")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=327536, Top:=1517946, _
    # TODO: Width:=5848931, Height:=3826050)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=6096000, Top:=1504075, _
    # TODO: Width:=5918244, Height:=5217400)
    #  ----- 슬라이드 16 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(16, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.Experiments")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=5918860, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = 4
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("We used the ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("MathBridge")
    runRange = paraRange.Characters
    runRange.Font.Name = "Courier New"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" dataset which is publicly available dataset containing ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("English descriptions of mathematical expressions and their corresponding LaTeX formats")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = 4
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Fine-tuning")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    paraRange.InsertAfter(" T5-large model")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = 4
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Using human evaluation as ground truth")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(255, 0, 0)
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.1 Dataset and setup")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=6859994, Top:=692304, _
    # TODO: Width:=5181583, Height:=5473392)
    #  ----- 슬라이드 17 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(17, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.Experiments")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838199, Top:=1825624, _
    # TODO: Width:=10336481, Height:=4530725)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("The ideal metric is human evaluation")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Conducted evaluations on the test dataset using ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("two groups ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("of human evaluators, H1 and H2, each consisting of ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("five members.")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("They were asked to score the predicted LaTeX format compared to the reference LaTeX format on a scale of 1 to 5")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1: “very inaccurate” 2: “inaccurate” 3: “neutral” 4: “accurate” 5:”very accurate”")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Pearson correlation coefficient and Spearman’s rank correlations coefficient")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.2 Human")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    paraRange.InsertAfter("evaluation")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  ----- 슬라이드 18 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(18, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.Experiments")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.3 Results")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=554473, Top:=1750344, _
    # TODO: Width:=11083054, Height:=3357312)
    #  ----- 슬라이드 19 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(19, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.Experiments")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Ablation study")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=864366, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = 3
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3.3 Results")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=2451153, Top:=2650468, _
    # TODO: Width:=7761625, Height:=3526495)
    #  ----- 슬라이드 20 추가 (제목 및 내용) -----
    ppSlide = ppPres.Slides.Add(20, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=365125, _
    # TODO: Width:=10515600, Height:=639709)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("4.Conclusion")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 내용 개체 틀 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=838200, Top:=1825625, _
    # TODO: Width:=10515600, Height:=4351338)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("We proposed ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(", an automatic metric for evaluating the accuracy of LaTeX formats")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" showed a ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("significantly higher correlation with human evaluation metrics than other metrics")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("One limitation of this study is that there is ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("no metric to verify whether the LaTeX format, when input into a compiler, can generate an error-free equation image.")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(255, 0, 0)
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("This is ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("a complex issue ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("because different ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("TeX")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" command sets and compilers can result in different compile errors. Further research should address this issue.")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=8610600, Top:=6356350, _
    # TODO: Width:=2743200, Height:=365125)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("/30")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  ----- 슬라이드 21 추가 (제목 슬라이드) -----
    ppSlide = ppPres.Slides.Add(21, ppLayoutTitleAndContent)
    #  텍스트 도형 추가: 제목 1
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=1160584, Top:=884855, _
    # TODO: Width:=9870831, Height:=2387600)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("TeXBLEU")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(": Automatic Metric for Evaluate LaTeX Format")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: 부제목 2
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=1956079, Top:=3364568, _
    # TODO: Width:=8279842, Height:=1655762)
    # TODO: Set textRange = shp.TextFrame.TextRange
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Kyudan")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Jung")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Nam-Joon Kim")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2∗")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Hyun Gon Ryu")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("3")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Sieun")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Hyeon")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Seung-")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("jun")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" Lee")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter(" , Hyuk-jae Lee")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("1")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Chung-Ang University, Seoul, Korea")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("2")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Seoul National ")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("Univeristy")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter(", Seoul, Korea")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("3")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    paraRange.InsertAfter("NVIDIA")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=0, Top:=5270359, _
    # TODO: Width:=1969477, Height:=1587641)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=203977, Top:=260510, _
    # TODO: Width:=1561521, Height:=890067)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=10578549, Top:=170309, _
    # TODO: Width:=1409474, Height:=1042255)
    #  이미지 플레이스홀더 추가
    #  실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: FileName:="C:\Path\To\Image.jpg", _
    # TODO: LinkToFile:=False, SaveWithDocument:=True, _
    # TODO: Left:=9334500, Top:=5649971, _
    # TODO: Width:=2857500, Height:=1208029)
    #  텍스트 도형 추가: TextBox 5
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=4655828, Top:=5842391, _
    # TODO: Width:=2893742, Height:=646331)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = 0
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = 2
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Presenter: Kyudan Jung")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    textRange.InsertAfter("\r\n")
    paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = 2
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("wjdrbeks1021@gmail.com")
    runRange = paraRange.Characters
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = 0
    #  텍스트 도형 추가: TextBox 3
    # Shapes.Add() 메서드는 파워포인트에서 지원하지 않음
    # 대신 AddShape() 메서드를 사용해야 함
    shp = ppSlide.Shapes.AddShape(1, 0, 0, 100, 100)  # 기본 사각형 추가
    # TODO: Left:=4233244, Top:=558499, _
    # TODO: Width:=3738909, Height:=830997)
    # TODO: Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = 0
    # TextFrame.AutoSize 속성 직접 설정하지 않음 - 대신 configure_text_frame 함수 사용
    configure_text_frame(shp)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter("Thank you ")
    runRange = paraRange.Characters
    runRange.Font.Size = 48.0
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    paraRange.InsertAfter("")
    runRange = paraRange.Characters
    runRange.Font.Size = 48.0
    runRange.Font.Bold = -1
    runRange.Font.Underline = 0
    # TODO: runRange.Font.Color.RGB = RGB(0, 51, 204)
    #  프레젠테이션 저장
    # TODO: ppPres.SaveAs "C:\Path\To\ReconstructedPresentation.pptx"
    #  객체 해제
    # TODO: Set runRange = Nothing
    # TODO: Set paraRange = Nothing
    # TODO: Set textRange = Nothing
    # TODO: Set tableCell = Nothing
    # TODO: Set shp = Nothing
    # TODO: Set ppSlide = Nothing
    # TODO: Set ppPres = Nothing
    # TODO: Set ppApp = Nothing
    print("프레젠테이션이 성공적으로 생성되었습니다.")

    try:
        print("프레젠테이션 저장 중...")

        # 임시 디렉토리에 저장 경로 설정
        temp_dir = tempfile.gettempdir()
        temp_filename = f"CreatePresentationFromAnalysis_{int(time.time())}.pptx"
        save_path = os.path.join(temp_dir, temp_filename)

        # 프레젠테이션 저장
        ppPres.SaveAs(save_path)
        print(f"프레젠테이션이 다음 위치에 저장되었습니다: {save_path}")

        return ppPres, save_path
    except Exception as e:
        print(f"저장 중 오류 발생: {str(e)}")

        # 대안으로 사용자에게 직접 저장하도록 안내
        print("PowerPoint 애플리케이션이 열려 있습니다. 직접 [파일 > 다른 이름으로 저장]을 선택하여 저장하세요.")

        return ppPres, None
if __name__ == "__main__":
    try:
        # 실행 경로 설정
        save_path = os.path.join(os.path.expanduser("~"), "Desktop", "CreatePresentationFromAnalysis.pptx")

        # 프레젠테이션 생성
        pres, saved_path = CreatePresentationFromAnalysis(save_path)

        if saved_path:
            print(f"프레젠테이션이 다음 위치에 저장되었습니다: {saved_path}")
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        print("상세 오류 정보:")
        import traceback
        traceback.print_exc()