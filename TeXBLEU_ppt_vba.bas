Sub CreatePresentationFromAnalysis()
    ' PowerPoint 애플리케이션 객체 생성
    Dim ppApp As Object
    Dim ppPres As Object
    Dim ppSlide As Object
    Dim shp As Object
    Dim textRange As Object
    Dim paraRange As Object
    Dim runRange As Object
    Dim tableCell As Object
    
    ' PowerPoint 상수 선언
    Const msoTrue As Long = -1
    Const msoFalse As Long = 0
    
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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ": Automatic Metric for Evaluate LaTeX Format"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(": Automatic Metric for Evaluate LaTeX Format") + 1, Len(": Automatic Metric for Evaluate LaTeX Format"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 부제목 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1956079, Top:=3364568, _
        Width:=8279842, Height:=1655762)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Kyudan"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Kyudan") + 1, Len("Kyudan"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Jung"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Jung") + 1, Len(" Jung"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Nam-Joon Kim"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Nam-Joon Kim") + 1, Len(" , Nam-Joon Kim"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2∗"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2∗") + 1, Len("2∗"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Hyun Gon Ryu"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Hyun Gon Ryu") + 1, Len(" , Hyun Gon Ryu"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "3"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3") + 1, Len("3"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , ") + 1, Len(" , "))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Sieun"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Sieun") + 1, Len("Sieun"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Hyeon"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Hyeon") + 1, Len(" Hyeon"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Seung-"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Seung-") + 1, Len(" , Seung-"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "jun"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("jun") + 1, Len("jun"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Lee"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Lee") + 1, Len(" Lee"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Hyuk-jae Lee"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Hyuk-jae Lee") + 1, Len(" , Hyuk-jae Lee"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Chung-Ang University, Seoul, Korea"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Chung-Ang University, Seoul, Korea") + 1, Len("Chung-Ang University, Seoul, Korea"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Seoul National "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Seoul National ") + 1, Len("Seoul National "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Univeristy"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Univeristy") + 1, Len("Univeristy"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ", Seoul, Korea"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(", Seoul, Korea") + 1, Len(", Seoul, Korea"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3") + 1, Len("3"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "NVIDIA"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("NVIDIA") + 1, Len("NVIDIA"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = msoFalse
    shp.TextFrame.AutoSize = SHAPE_TO_FIT_TEXT (1)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = CENTER (2)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "wjdrbeks1021@gmail.com"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("wjdrbeks1021@gmail.com") + 1, Len("wjdrbeks1021@gmail.com"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 2 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(2, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=1325563)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Overview"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Overview") + 1, Len("Overview"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1527349, _
        Width:=10515600, Height:=4965526)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Problem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Problem Define") + 1, Len("Problem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "1.1 Metric for LaTeX format"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.1 Metric for LaTeX format") + 1, Len("1.1 Metric for LaTeX format"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "1.2 Limitations of existing metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.2 Limitations of existing metrics") + 1, Len("1.2 Limitations of existing metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Method") + 1, Len("Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "2.1 Tokenizer trained by "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.1 Tokenizer trained by ") + 1, Len("2.1 Tokenizer trained by "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "arXiv"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("arXiv") + 1, Len("arXiv"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " papers"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" papers") + 1, Len(" papers"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "2.2 Fine-tuned embedding model"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.2 Fine-tuned embedding model") + 1, Len("2.2 Fine-tuned embedding model"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "2.3 Computing "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing ") + 1, Len("2.3 Computing "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" ") + 1, Len(" "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("argorithm") + 1, Len("argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Experiments"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Experiments") + 1, Len("Experiments"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "3.1 Dataset and setup"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.1 Dataset and setup") + 1, Len("3.1 Dataset and setup"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "3.2 Human evaluation"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.2 Human evaluation") + 1, Len("3.2 Human evaluation"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 1
    paraRange.InsertAfter "3.3 Result"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.3 Result") + 1, Len("3.3 Result"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Conclusion"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Conclusion") + 1, Len("Conclusion"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 3
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 3 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(3, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.Problem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.Problem Define") + 1, Len("1.Problem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838199, Top:=1825625, _
        Width:=10930247, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "LaTeX format"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("LaTeX format") + 1, Len("LaTeX format"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " : Useful format in Mathematics, Computer Science, etc."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" : Useful format in Mathematics, Computer Science, etc.") + 1, Len(" : Useful format in Mathematics, Computer Science, etc."))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Example: In converting spoken mathematical expressions into LaTeX"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Example: In converting spoken mathematical expressions into LaTeX") + 1, Len("Example: In converting spoken mathematical expressions into LaTeX"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Command-based language such as C or SQL"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Command-based language such as C or SQL") + 1, Len("Command-based language such as C or SQL"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Suitable evaluation metric is required"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Suitable evaluation metric is required") + 1, Len("Suitable evaluation metric is required"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.1 Metric for LaTeX Format "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.1 Metric for LaTeX Format ") + 1, Len("1.1 Metric for LaTeX Format "))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 4 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(4, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.Promblem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.Promblem Define") + 1, Len("1.Promblem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "BLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("BLEU") + 1, Len("BLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.2 Limitations of existing metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.2 Limitations of existing metrics") + 1, Len("1.2 Limitations of existing metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.Promblem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.Promblem Define") + 1, Len("1.Promblem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "sacreBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("sacreBLEU") + 1, Len("sacreBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.2 Limitations of existing metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.2 Limitations of existing metrics") + 1, Len("1.2 Limitations of existing metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.Promblem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.Promblem Define") + 1, Len("1.Promblem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Character Error Rate (CER)"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Character Error Rate (CER)") + 1, Len("Character Error Rate (CER)"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.2 Limitations of existing metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.2 Limitations of existing metrics") + 1, Len("1.2 Limitations of existing metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.Promblem Define"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.Promblem Define") + 1, Len("1.Promblem Define"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Word Error Rate (WER)"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Word Error Rate (WER)") + 1, Len("Word Error Rate (WER)"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1.2 Limitations of existing metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1.2 Limitations of existing metrics") + 1, Len("1.2 Limitations of existing metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Requiring "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Requiring ") + 1, Len("Requiring "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "LaTeX-specified tokenizer"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("LaTeX-specified tokenizer") + 1, Len("LaTeX-specified tokenizer"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Collecting papers in "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Collecting papers in ") + 1, Len("Collecting papers in "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "arXiv"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("arXiv") + 1, Len("arXiv"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " ."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" .") + 1, Len(" ."))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "tex"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("tex") + 1, Len("tex"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " files"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" files") + 1, Len(" files"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Byte-pair encoding (BPE) based tokenizer"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Byte-pair encoding (BPE) based tokenizer") + 1, Len("Byte-pair encoding (BPE) based tokenizer"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Unlike other tokenizers, the tokenizer is designed to "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Unlike other tokenizers, the tokenizer is designed to ") + 1, Len("Unlike other tokenizers, the tokenizer is designed to "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "capture LaTeX grammar elements"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("capture LaTeX grammar elements") + 1, Len("capture LaTeX grammar elements"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " and tokenize them in a way that "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" and tokenize them in a way that ") + 1, Len(" and tokenize them in a way that "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "reflects LaTeX’s unique structure"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("reflects LaTeX’s unique structure") + 1, Len("reflects LaTeX’s unique structure"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.1 Tokenizer trained by arXiv papers"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.1 Tokenizer trained by arXiv papers") + 1, Len("2.1 Tokenizer trained by arXiv papers"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 9 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(9, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=371809, Top:=4283565, _
        Width:=11634143, Height:=1893397)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = JUSTIFY (4)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Fig. 4. The results of tokenizing the quadratic formula in LaTeX format using various models. "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Fig. 4. The results of tokenizing the quadratic formula in LaTeX format using various models. ") + 1, Len("Fig. 4. The results of tokenizing the quadratic formula in LaTeX format using various models. "))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "The "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("The ") + 1, Len("The "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "green boxes indicate sections where LaTeX commands were successfully tokenized as complete chunks"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("green boxes indicate sections where LaTeX commands were successfully tokenized as complete chunks") + 1, Len("green boxes indicate sections where LaTeX commands were successfully tokenized as complete chunks"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ", while the "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(", while the ") + 1, Len(", while the "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "red boxes represent sections where the LaTeX commands were fragmented, losing their meaning during tokenization"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("red boxes represent sections where the LaTeX commands were fragmented, losing their meaning during tokenization") + 1, Len("red boxes represent sections where the LaTeX commands were fragmented, losing their meaning during tokenization"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(255, 0, 0)
    paraRange.InsertAfter ". It is evident that the tokenizer we developed using the "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(". It is evident that the tokenizer we developed using the ") + 1, Len(". It is evident that the tokenizer we developed using the "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "arXiv"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("arXiv") + 1, Len("arXiv"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " paper corpus performs the best in accurately tokenizing LaTeX commands."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" paper corpus performs the best in accurately tokenizing LaTeX commands.") + 1, Len(" paper corpus performs the best in accurately tokenizing LaTeX commands."))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.1 Tokenizer trained by arXiv papers"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.1 Tokenizer trained by arXiv papers") + 1, Len("2.1 Tokenizer trained by arXiv papers"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Based on new tokenizer, we made "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Based on new tokenizer, we made ") + 1, Len("Based on new tokenizer, we made "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "new embedding model"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("new embedding model") + 1, Len("new embedding model"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Fine-tuning a publicly available pretrained "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Fine-tuning a publicly available pretrained ") + 1, Len("Fine-tuning a publicly available pretrained "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "GPT-2 embedding model"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("GPT-2 embedding model") + 1, Len("GPT-2 embedding model"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.2 Fine-tuned embedding model"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.2 Fine-tuned embedding model") + 1, Len("2.2 Fine-tuned embedding model"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 11 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(11, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.3 Computing TeXBLEU argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing TeXBLEU argorithm") + 1, Len("2.3 Computing TeXBLEU argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.3 Computing TeXBLEU argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing TeXBLEU argorithm") + 1, Len("2.3 Computing TeXBLEU argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.3 Computing TeXBLEU argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing TeXBLEU argorithm") + 1, Len("2.3 Computing TeXBLEU argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825624, _
        Width:=10515600, Height:=4329699)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Finally, "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Finally, ") + 1, Len("Finally, "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " is finally equal to the following expression"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" is finally equal to the following expression") + 1, Len(" is finally equal to the following expression"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.3 Computing TeXBLEU argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing TeXBLEU argorithm") + 1, Len("2.3 Computing TeXBLEU argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.Method"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.Method") + 1, Len("2.Method"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2.3 Computing TeXBLEU argorithm"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2.3 Computing TeXBLEU argorithm") + 1, Len("2.3 Computing TeXBLEU argorithm"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.Experiments"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.Experiments") + 1, Len("3.Experiments"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=5918860, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = JUSTIFY (4)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "We used the "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("We used the ") + 1, Len("We used the "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "MathBridge"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("MathBridge") + 1, Len("MathBridge"))
    runRange.Font.Name = "Courier New"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " dataset which is publicly available dataset containing "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" dataset which is publicly available dataset containing ") + 1, Len(" dataset which is publicly available dataset containing "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "English descriptions of mathematical expressions and their corresponding LaTeX formats"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("English descriptions of mathematical expressions and their corresponding LaTeX formats") + 1, Len("English descriptions of mathematical expressions and their corresponding LaTeX formats"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = JUSTIFY (4)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Fine-tuning"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Fine-tuning") + 1, Len("Fine-tuning"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    paraRange.InsertAfter " T5-large model"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" T5-large model") + 1, Len(" T5-large model"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = JUSTIFY (4)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Using human evaluation as ground truth"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Using human evaluation as ground truth") + 1, Len("Using human evaluation as ground truth"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(255, 0, 0)

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.1 Dataset and setup"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.1 Dataset and setup") + 1, Len("3.1 Dataset and setup"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.Experiments"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.Experiments") + 1, Len("3.Experiments"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838199, Top:=1825624, _
        Width:=10336481, Height:=4530725)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "The ideal metric is human evaluation"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("The ideal metric is human evaluation") + 1, Len("The ideal metric is human evaluation"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Conducted evaluations on the test dataset using "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Conducted evaluations on the test dataset using ") + 1, Len("Conducted evaluations on the test dataset using "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "two groups "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("two groups ") + 1, Len("two groups "))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "of human evaluators, H1 and H2, each consisting of "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("of human evaluators, H1 and H2, each consisting of ") + 1, Len("of human evaluators, H1 and H2, each consisting of "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "five members."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("five members.") + 1, Len("five members."))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "They were asked to score the predicted LaTeX format compared to the reference LaTeX format on a scale of 1 to 5"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("They were asked to score the predicted LaTeX format compared to the reference LaTeX format on a scale of 1 to 5") + 1, Len("They were asked to score the predicted LaTeX format compared to the reference LaTeX format on a scale of 1 to 5"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1: “very inaccurate” 2: “inaccurate” 3: “neutral” 4: “accurate” 5:”very accurate”"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1: “very inaccurate” 2: “inaccurate” 3: “neutral” 4: “accurate” 5:”very accurate”") + 1, Len("1: “very inaccurate” 2: “inaccurate” 3: “neutral” 4: “accurate” 5:”very accurate”"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Pearson correlation coefficient and Spearman’s rank correlations coefficient"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Pearson correlation coefficient and Spearman’s rank correlations coefficient") + 1, Len("Pearson correlation coefficient and Spearman’s rank correlations coefficient"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.2 Human"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.2 Human") + 1, Len("3.2 Human"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" ") + 1, Len(" "))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "evaluation"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("evaluation") + 1, Len("evaluation"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 18 추가 (제목 및 내용) -----
    Set ppSlide = ppPres.Slides.Add(18, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=365125, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.Experiments"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.Experiments") + 1, Len("3.Experiments"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.3 Results"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.3 Results") + 1, Len("3.3 Results"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.Experiments"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.Experiments") + 1, Len("3.Experiments"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Ablation study"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Ablation study") + 1, Len("Ablation study"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=864366, _
        Width:=10515600, Height:=639709)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.VerticalAnchor = MIDDLE (3)
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3.3 Results"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3.3 Results") + 1, Len("3.3 Results"))
    runRange.Font.Name = "Arial"
    runRange.Font.Size = 32.0
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = NONE (0)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "4.Conclusion"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("4.Conclusion") + 1, Len("4.Conclusion"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 내용 개체 틀 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=838200, Top:=1825625, _
        Width:=10515600, Height:=4351338)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "We proposed "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("We proposed ") + 1, Len("We proposed "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ", an automatic metric for evaluating the accuracy of LaTeX formats"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(", an automatic metric for evaluating the accuracy of LaTeX formats") + 1, Len(", an automatic metric for evaluating the accuracy of LaTeX formats"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " showed a "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" showed a ") + 1, Len(" showed a "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "significantly higher correlation with human evaluation metrics than other metrics"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("significantly higher correlation with human evaluation metrics than other metrics") + 1, Len("significantly higher correlation with human evaluation metrics than other metrics"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "One limitation of this study is that there is "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("One limitation of this study is that there is ") + 1, Len("One limitation of this study is that there is "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "no metric to verify whether the LaTeX format, when input into a compiler, can generate an error-free equation image."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("no metric to verify whether the LaTeX format, when input into a compiler, can generate an error-free equation image.") + 1, Len("no metric to verify whether the LaTeX format, when input into a compiler, can generate an error-free equation image."))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(255, 0, 0)
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "This is "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("This is ") + 1, Len("This is "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "a complex issue "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("a complex issue ") + 1, Len("a complex issue "))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "because different "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("because different ") + 1, Len("because different "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "TeX"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeX") + 1, Len("TeX"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " command sets and compilers can result in different compile errors. Further research should address this issue."
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" command sets and compilers can result in different compile errors. Further research should address this issue.") + 1, Len(" command sets and compilers can result in different compile errors. Further research should address this issue."))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 슬라이드 번호 개체 틀 4
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=8610600, Top:=6356350, _
        Width:=2743200, Height:=365125)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "/30"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("/30") + 1, Len("/30"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' ----- 슬라이드 21 추가 (제목 슬라이드) -----
    Set ppSlide = ppPres.Slides.Add(21, 7)

    ' 텍스트 도형 추가: 제목 1
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1160584, Top:=884855, _
        Width:=9870831, Height:=2387600)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "TeXBLEU"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("TeXBLEU") + 1, Len("TeXBLEU"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ": Automatic Metric for Evaluate LaTeX Format"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(": Automatic Metric for Evaluate LaTeX Format") + 1, Len(": Automatic Metric for Evaluate LaTeX Format"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: 부제목 2
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=1956079, Top:=3364568, _
        Width:=8279842, Height:=1655762)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.AutoSize = TEXT_TO_FIT_SHAPE (2)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Kyudan"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Kyudan") + 1, Len("Kyudan"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Jung"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Jung") + 1, Len(" Jung"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Nam-Joon Kim"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Nam-Joon Kim") + 1, Len(" , Nam-Joon Kim"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2∗"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2∗") + 1, Len("2∗"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Hyun Gon Ryu"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Hyun Gon Ryu") + 1, Len(" , Hyun Gon Ryu"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "3"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3") + 1, Len("3"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , ") + 1, Len(" , "))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Sieun"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Sieun") + 1, Len("Sieun"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Hyeon"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Hyeon") + 1, Len(" Hyeon"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Seung-"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Seung-") + 1, Len(" , Seung-"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "jun"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("jun") + 1, Len("jun"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " Lee"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" Lee") + 1, Len(" Lee"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter " , Hyuk-jae Lee"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(" , Hyuk-jae Lee") + 1, Len(" , Hyuk-jae Lee"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "1"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("1") + 1, Len("1"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Chung-Ang University, Seoul, Korea"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Chung-Ang University, Seoul, Korea") + 1, Len("Chung-Ang University, Seoul, Korea"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "2"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("2") + 1, Len("2"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Seoul National "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Seoul National ") + 1, Len("Seoul National "))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "Univeristy"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Univeristy") + 1, Len("Univeristy"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter ", Seoul, Korea"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len(", Seoul, Korea") + 1, Len(", Seoul, Korea"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "3"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("3") + 1, Len("3"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    paraRange.InsertAfter "NVIDIA"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("NVIDIA") + 1, Len("NVIDIA"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

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
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = msoFalse
    shp.TextFrame.AutoSize = SHAPE_TO_FIT_TEXT (1)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.Alignment = CENTER (2)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Presenter: Kyudan Jung"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Presenter: Kyudan Jung") + 1, Len("Presenter: Kyudan Jung"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse
    textRange.InsertAfter vbCrLf
    Set paraRange = textRange.Paragraphs(textRange.Paragraphs.Count)
    paraRange.ParagraphFormat.Alignment = CENTER (2)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "wjdrbeks1021@gmail.com"
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("wjdrbeks1021@gmail.com") + 1, Len("wjdrbeks1021@gmail.com"))
    runRange.Font.Name = "Arial"
    runRange.Font.Underline = msoFalse

    ' 텍스트 도형 추가: TextBox 3
    Set shp = ppSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=4233244, Top:=558499, _
        Width:=3738909, Height:=830997)
    Set textRange = shp.TextFrame.TextRange
    shp.TextFrame.WordWrap = msoFalse
    shp.TextFrame.AutoSize = SHAPE_TO_FIT_TEXT (1)
    shp.TextFrame.Top = 45720
    shp.TextFrame.Bottom = 45720
    shp.TextFrame.Left = 91440
    shp.TextFrame.Right = 91440
    textRange.Text = ""
    Set paraRange = textRange.Paragraphs(1)
    paraRange.ParagraphFormat.IndentLevel = 0
    paraRange.InsertAfter "Thank you "
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("Thank you ") + 1, Len("Thank you "))
    runRange.Font.Size = 48.0
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)
    paraRange.InsertAfter ""
    Set runRange = paraRange.Characters(Len(paraRange.Text) - Len("") + 1, Len(""))
    runRange.Font.Size = 48.0
    runRange.Font.Bold = msoTrue
    runRange.Font.Underline = msoFalse
    runRange.Font.Color.RGB = RGB(0, 51, 204)

    ' 프레젠테이션 저장
    ppPres.SaveAs "C:\Path\To\ReconstructedPresentation.pptx"
    
    ' 객체 해제
    Set runRange = Nothing
    Set paraRange = Nothing
    Set textRange = Nothing
    Set tableCell = Nothing
    Set shp = Nothing
    Set ppSlide = Nothing
    Set ppPres = Nothing
    Set ppApp = Nothing
    
    MsgBox "프레젠테이션이 성공적으로 생성되었습니다."
End Sub