# 🚀 Talk-to-Your-Slides

<div align="center">
   
  **Talk to Your Slides: Real‑Time Agent‑Based PowerPoint Automation with Large Language Models**
  
🗒️ Our **research paper** is out!

https://arxiv.org/abs/2505.11604

🖥️ Our **code** is out!
</div>

## 📖 Overview

Our PPT Agent can modify PowerPoint presentations in real-time while PowerPoint is open.<br>
It receives natural language-based user commands and successfully modifies the PPT through interaction with the agent, presenting the updated PowerPoint to the user.<br><br>

> **News about the paper publication will be available first on [LinkedIn](https://www.linkedin.com/in/kyudanjung/) and [Research Blog](https://sites.google.com/view/kyudanjung/).**


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
@misc{jung2025talkslideslanguagedrivenagents,
      title={Talk to Your Slides: Language-Driven Agents for Efficient Slide Editing}, 
      author={Kyudan Jung and Hojun Cho and Jooyeol Yun and Soyoung Yang and Jaehyeok Jang and Jaegul Choo},
      year={2025},
      eprint={2505.11604},
      archivePrefix={arXiv},
      primaryClass={cs.CL},
      url={https://arxiv.org/abs/2505.11604}, 
}
```

