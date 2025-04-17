# 🚀 Talk-to-Your-Slides

<div align="center">
  
 
  **Talk to Your Slides: Real‑Time Agent‑Based PowerPoint Automation with Large Language Models**
  
  [![Stars](https://img.shields.io/github/stars/KyuDan1/PPTAgent-4.16?style=social)](https://github.com/KyuDan1/PPTAgent-4.16/stargazers)
  [![Demo](https://img.shields.io/badge/Demo-Watch%20Now-red)](https://youtu.be/your-demo-link)
  
</div>

## 📖 Overview

Our PPT Agent can modify PowerPoint presentations in real-time while PowerPoint is open.<br>
It receives natural language-based user commands and successfully modifies the PPT through interaction with the agent, presenting the updated PowerPoint to the user.<br><br>
**✨Our research paper on PPT Agent will be published soon!✨<br><br> It covers in-depth content about the latest AI-based presentation generation technology, so stay tuned.**

> **News about the paper publication will be available first on [LinkedIn](https://www.linkedin.com/in/kyudanjung/) and [Research Blog](https://sites.google.com/view/kyudanjung/).**


## ✨ Key Features

### Overall
<img src="fig1.png">

## 🎬 Demo Video

<div align="center">
  
  [![PPT Agent Demo Video](https://img.youtube.com/vi/your-video-id/0.jpg)](https://www.youtube.com/watch?v=your-video-id "PPT Agent Demo Video")
  
  [Watch the full demo on YouTube](https://youtu.be/your-demo-link)
  
</div>

## 🛠️ Installation Guide
### Recommended for Python in Windows.

### conda environment
```bash
pip install -r 'requirements.txt'
```
- Then make 'credentials.yml' on pptagent-4.16 directory.
you should make like below.
```yml
gpt-4.1-mini:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"
gpt-4.1-nano:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"
gemini-1.5-flash:
  api_key: "YOUR_GEMINI_API_KEY"
```
- add .env file in pptagent direction on pptagent-4.16/pptagent.
```bash
python pptagent/main.py
```

## 📊 How to Cite

If you use PPT Agent in your research or project, please cite as follows:

```bibtex
@software{ppt_agent2025,
  author = {Kyudan Jung and Hojun Cho and Jooyeol Yun and Jaegul Choo},
  title = {PPT Agent: AI-Powered Presentation Generator},
  url = {https://github.com/KyuDan1/PPTAgent-4.16},
  version = {1.0.0},
  year = {2025},
}
```

Or you can cite it briefly as:

```
Your Name. (2025). PPT Agent: AI-Powered Presentation Generator [Software]. Available from https://github.com/KyuDan1/PPTAgent-4.16
```