# 🚀 Talk-to-Your-Slides

<div align="center">
   
  **Talk to Your Slides: Real‑Time Agent‑Based PowerPoint Automation with Large Language Models**
  
🗒️ Our **research paper** on PPT Agent will be published soon!

🖥️ Our **code** will be out soon!
</div>

## 📖 Overview

Our PPT Agent can modify PowerPoint presentations in real-time while PowerPoint is open.<br>
It receives natural language-based user commands and successfully modifies the PPT through interaction with the agent, presenting the updated PowerPoint to the user.<br><br>
**✨Our research paper on PPT Agent will be published soon!✨<br><br> It covers in-depth content about the latest AI-based presentation generation technology, so stay tuned.**

> **News about the paper publication will be available first on [LinkedIn](https://www.linkedin.com/in/kyudanjung/) and [Research Blog](https://sites.google.com/view/kyudanjung/).**


## ✨ Key Features

### Overall


- **Figure 1. Diagram of the Talk-to-Your-Slides System.** This figure illustrates the overall workflow of the system. When a user inputs a command in natural language, the **Planner** first generates a plan. Then, the **Parser** parses PowerPoint-related information using a rule-based approach. The **Processor** interprets this parsed information, and the **Applier** generates Python code to apply the intended changes to the PowerPoint file. Finally, the **Reporter** summarizes and reports all executed commands back to the user.

## 🎬 Demo Videos

<div align="center">

[![CamelCase Demo](https://img.youtube.com/vi/9nJ0-yofr7Y/0.jpg)](https://youtu.be/9nJ0-yofr7Y "CamelCase Formatting")  
**CamelCase**  
*Demo prompt:* “Please update all English on ppt slides number 7 to camelCase formatting.”  

[![Only English → Blue](https://img.youtube.com/vi/eVSs6xi-bEs/0.jpg)](https://youtu.be/eVSs6xi-bEs "Only English Blue")  
**Only English → Blue**  
*Demo prompt:* “Please change only English into blue color in slide number 3.”  

[![Typo Checking Demo](https://img.youtube.com/vi/rBIBsnWX3W0/0.jpg)](https://youtu.be/rBIBsnWX3W0 "Typo Checking & Correction")  
**Typo Checking & Correction**  
*Demo prompt:* “Please check ppt slides number 4 for any typos or errors, correct them.”  

[![Translate to English](https://img.youtube.com/vi/GLS_9xh2C-4/0.jpg)](https://youtu.be/GLS_9xh2C-4 "Translate Slides")  
**Translate to English**  
*Demo prompt:* “Please translate ppt slides number 5 into English.”  

[![Slide‑Notes Script](https://img.youtube.com/vi/5vzYd5ov_Cs/0.jpg)](https://youtu.be/5vzYd5ov_Cs "Generate Slide Notes")  
**Slide‑Notes Script**  
*Demo prompt:* “Please create a full script for ppt slides number 3 and add the script to the slide notes.”  

</div>

## 🛠️ Installation Guide
### Recommended for Python in Windows.

### conda environment
```bash
pip install -r 'requirements.txt'
```
- Then make 'credentials.yml' on (will be out soon) directory.
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
- add .env file in pptagent direction on (will be out soon)/pptagent.
```bash
python pptagent/main.py
```

## 📊 How to Cite

If you use PPT Agent in your research or project, please cite as follows:

```bibtex
@software{ppt_agent2025,
  author = {Kyudan Jung and Hojun Cho and Jooyeol Yun and Jaegul Choo},
  title = {PPT Agent: AI-Powered Presentation Generator},
  url = {https://github.com/KyuDan1/(will be out soon)},
  version = {1.0.0},
  year = {2025},
}
```

Or you can cite it briefly as:

```
Your Name. (2025). PPT Agent: AI-Powered Presentation Generator [Software]. Available from https://github.com/KyuDan1/(will be out soon)
```
