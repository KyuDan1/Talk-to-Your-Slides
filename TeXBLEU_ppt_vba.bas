Sub CreatePresentationFromAnalysis()
    ' PowerPoint 애플리케이션 객체 생성
    Dim ppApp As Object
    Dim ppPres As Object
    Dim ppSlide As Object
    Dim shp As Object
    
    ' PowerPoint 시작
    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = True ' PowerPoint 창을 보이게 설정
    
    ' 새 프레젠테이션 생성
    Set ppPres = ppApp.Presentations.Add
    
    ' 슬라이드 크기 설정
    ppPres.PageSetup.SlideWidth = 12192000
    ppPres.PageSetup.SlideHeight = 6858000

    ' ----- 슬라이드 1 추가 (제목 슬라이드) -----
    Set ppSlide = ppPres.Slides.Add(1, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1160584, Top:=884855, _
        Width:=9870831, Height:=2387600)
    shp.TextFrame.TextRange.Text = "TeXBLEU: Automatic Metric for Evaluate LaTeX Format"

    ' 텍스트 도형 추가: 부제목 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1956079, Top:=3364568, _
        Width:=8279842, Height:=1655762)
    shp.TextFrame.TextRange.Text = "Kyudan Jung1 , Nam-Joon Kim2∗ , Hyun Gon Ryu3 , Sieun Hyeon2 , Seung-jun Lee1 , Hyuk-jae Lee2" & Chr(13) & "1Chung-Ang University, Seoul, Korea" & Chr(13) & "2Seoul National Univeristy, Seoul, Korea" & Chr(13) & "3NVIDIA"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=0, Top:=5270359, _
        Width:=1969477, Height:=1587641)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=203977, Top:=260510, _
        Width:=1561521, Height:=890067)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=10578549, Top:=170309, _
        Width:=1409474, Height:=1042255)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=9334500, Top:=5649971, _
        Width:=2857500, Height:=1208029)

    ' 텍스트 도형 추가: TextBox 5
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=4655828, Top:=5842391, _
        Width:=2893741, Height:=369332)
    shp.TextFrame.TextRange.Text = "wjdrbeks1021@gmail.com"

    ' ----- 슬라이드 2 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(2, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=1325563)
    shp.TextFrame.TextRange.Text = "Overview"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1527349, _
        Width:=10515600, Height:=4965526)
    shp.TextFrame.TextRange.Text = "Problem Define" & Chr(13) & "1.1 Metric for LaTeX format" & Chr(13) & "1.2 Limitations of existing metrics" & Chr(13) & "Method" & Chr(13) & "2.1 Tokenizer trained by arXiv papers" & Chr(13) & "2.2 Fine-tuned embedding model" & Chr(13) & "2.3 Computing TeXBLEU argorithm" & Chr(13) & "Experiments" & Chr(13) & "3.1 Dataset and setup" & Chr(13) & "3.2 Human evaluation" & Chr(13) & "3.3 Result" & Chr(13) & "Conclusion"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 3
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "2/30"

    ' ----- 슬라이드 3 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(3, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.Problem Define"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838199, Top:=1825625, _
        Width:=10930247, Height:=4351338)
    shp.TextFrame.TextRange.Text = "LaTeX format : Useful format in Mathematics, Computer Science, etc." & Chr(13) & "Example: In converting spoken mathematical expressions into LaTeX" & Chr(13) & "Command-based language such as C or SQL" & Chr(13) & "Suitable evaluation metric is required"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "3/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.1 Metric for LaTeX Format"

    ' ----- 슬라이드 4 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(4, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.Promblem Define"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "BLEU"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "4/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.2 Limitations of existing metrics"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2381459, Top:=2744905, _
        Width:=7429081, Height:=2512778)

    ' ----- 슬라이드 5 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(5, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.Promblem Define"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "sacreBLEU"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "5/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.2 Limitations of existing metrics"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2424350, Top:=2823116, _
        Width:=7343299, Height:=2651408)

    ' ----- 슬라이드 6 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(6, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.Promblem Define"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "Character Error Rate (CER)"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "6/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.2 Limitations of existing metrics"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2596585, Top:=2897668, _
        Width:=6998829, Height:=2340707)

    ' ----- 슬라이드 7 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(7, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.Promblem Define"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "Word Error Rate (WER)"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "7/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "1.2 Limitations of existing metrics"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2338720, Top:=2908720, _
        Width:=7514560, Height:=2542054)

    ' ----- 슬라이드 8 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(8, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "Requiring LaTeX-specified tokenizer" & Chr(13) & "Collecting papers in arXiv .tex files" & Chr(13) & "Byte-pair encoding (BPE) based tokenizer" & Chr(13) & "Unlike other tokenizers, the tokenizer is designed to capture LaTeX grammar elements and tokenize them in a way that reflects LaTeX’s unique structure"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "8/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.1 Tokenizer trained by arXiv papers"

    ' ----- 슬라이드 9 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(9, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=371809, Top:=4283565, _
        Width:=11634143, Height:=1893397)
    shp.TextFrame.TextRange.Text = "Fig. 4. The results of tokenizing the quadratic formula in LaTeX format using various models. The green boxes indicate sections where LaTeX commands were successfully tokenized as complete chunks, while the red boxes represent sections where the LaTeX commands were fragmented, losing their meaning during tokenization. It is evident that the tokenizer we developed using the arXiv paper corpus performs the best in accurately tokenizing LaTeX commands."

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "9/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.1 Tokenizer trained by arXiv papers"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=647972, Top:=1908518, _
        Width:=10981990, Height:=2198512)

    ' ----- 슬라이드 10 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(10, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "Based on new tokenizer, we made new embedding model" & Chr(13) & "Fine-tuning a publicly available pretrained GPT-2 embedding model"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "10/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.2 Fine-tuned embedding model"

    ' ----- 슬라이드 11 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(11, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "11/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.3 Computing TeXBLEU argorithm"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2461240, Top:=2358845, _
        Width:=7520960, Height:=927651)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2960064, Top:=5504189, _
        Width:=6920207, Height:=633833)

    ' ----- 슬라이드 12 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(12, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "12/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.3 Computing TeXBLEU argorithm"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2495537, Top:=1504075, _
        Width:=6719715, Height:=5345045)

    ' ----- 슬라이드 13 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(13, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "13/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.3 Computing TeXBLEU argorithm"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2730326, Top:=4219532, _
        Width:=6493842, Height:=1061213)

    ' ----- 슬라이드 14 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(14, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825624, _
        Width:=10515600, Height:=4329699)
    shp.TextFrame.TextRange.Text = "Finally, TeXBLEU is finally equal to the following expression"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "14/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.3 Computing TeXBLEU argorithm"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2512660, Top:=3230252, _
        Width:=7166679, Height:=1265938)

    ' ----- 슬라이드 15 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(15, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.Method"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "15/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "2.3 Computing TeXBLEU argorithm"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=327536, Top:=1517946, _
        Width:=5848931, Height:=3826050)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=6096000, Top:=1504075, _
        Width:=5918244, Height:=5217400)

    ' ----- 슬라이드 16 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(16, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.Experiments"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=5918860, Height:=4351338)
    shp.TextFrame.TextRange.Text = "We used the MathBridge dataset which is publicly available dataset containing English descriptions of mathematical expressions and their corresponding LaTeX formats" & Chr(13) & "Fine-tuning T5-large model" & Chr(13) & "Using human evaluation as ground truth"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "16/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.1 Dataset and setup"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=6859994, Top:=692304, _
        Width:=5181583, Height:=5473392)

    ' ----- 슬라이드 17 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(17, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.Experiments"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838199, Top:=1825624, _
        Width:=10336481, Height:=4530725)
    shp.TextFrame.TextRange.Text = "The ideal metric is human evaluation" & Chr(13) & "Conducted evaluations on the test dataset using two groups of human evaluators, H1 and H2, each consisting of five members." & Chr(13) & "They were asked to score the predicted LaTeX format compared to the reference LaTeX format on a scale of 1 to 5" & Chr(13) & "1: “very inaccurate” 2: “inaccurate” 3: “neutral” 4: “accurate” 5:”very accurate”" & Chr(13) & "Pearson correlation coefficient and Spearman’s rank correlations coefficient"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "17/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.2 Human evaluation"

    ' ----- 슬라이드 18 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(18, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.Experiments"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = ""

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "18/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.3 Results"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=554473, Top:=1750344, _
        Width:=11083054, Height:=3357312)

    ' ----- 슬라이드 19 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(19, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.Experiments"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "Ablation study"

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "19/30"

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "3.3 Results"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=2451153, Top:=2650468, _
        Width:=7761625, Height:=3526495)

    ' ----- 슬라이드 20 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(20, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    shp.TextFrame.TextRange.Text = "4.Conclusion"

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    shp.TextFrame.TextRange.Text = "We proposed TeXBLEU, an automatic metric for evaluating the accuracy of LaTeX formats" & Chr(13) & "TeXBLEU showed a significantly higher correlation with human evaluation metrics than other metrics" & Chr(13) & "One limitation of this study is that there is no metric to verify whether the LaTeX format, when input into a compiler, can generate an error-free equation image." & Chr(13) & "This is a complex issue because different TeX command sets and compilers can result in different compile errors. Further research should address this issue."

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    shp.TextFrame.TextRange.Text = "20/30"

    ' ----- 슬라이드 21 추가 (제목 슬라이드) -----
    Set ppSlide = ppPres.Slides.Add(21, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1160584, Top:=884855, _
        Width:=9870831, Height:=2387600)
    shp.TextFrame.TextRange.Text = "TeXBLEU: Automatic Metric for Evaluate LaTeX Format"

    ' 텍스트 도형 추가: 부제목 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1956079, Top:=3364568, _
        Width:=8279842, Height:=1655762)
    shp.TextFrame.TextRange.Text = "Kyudan Jung1 , Nam-Joon Kim2∗ , Hyun Gon Ryu3 , Sieun Hyeon2 , Seung-jun Lee1 , Hyuk-jae Lee2" & Chr(13) & "1Chung-Ang University, Seoul, Korea" & Chr(13) & "2Seoul National Univeristy, Seoul, Korea" & Chr(13) & "3NVIDIA"

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=0, Top:=5270359, _
        Width:=1969477, Height:=1587641)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=203977, Top:=260510, _
        Width:=1561521, Height:=890067)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=10578549, Top:=170309, _
        Width:=1409474, Height:=1042255)

    ' 이미지 플레이스홀더 추가
    ' 실제 이미지를 추가하려면 이미지 파일 경로 지정 필요
    Set shp = ppSlide.Shapes.AddPicture( _
        FileName:="C:\Path\To\Image.jpg", _
        LinkToFile:=False, SaveWithDocument:=True, _
        Left:=9334500, Top:=5649971, _
        Width:=2857500, Height:=1208029)

    ' 텍스트 도형 추가: TextBox 5
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=4655828, Top:=5842391, _
        Width:=2893742, Height:=646331)
    shp.TextFrame.TextRange.Text = "Presenter: Kyudan Jung" & Chr(13) & "wjdrbeks1021@gmail.com"

    ' 텍스트 도형 추가: TextBox 3
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=4233244, Top:=558499, _
        Width:=3738909, Height:=830997)
    shp.TextFrame.TextRange.Text = "Thank you "

    ' 프레젠테이션 저장
    ppPres.SaveAs "C:\Path\To\ReconstructedPresentation.pptx"
    
    ' 객체 해제
    Set ppSlide = Nothing
    Set ppPres = Nothing
    Set ppApp = Nothing
    
    MsgBox "프레젠테이션이 성공적으로 생성되었습니다."
End Sub