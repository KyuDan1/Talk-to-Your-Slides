# 🚀 Talk-to-Your-Slides [![Stars](https://img.shields.io/github/stars/KyuDan1/Talk-to-Your-Slides?style=social)](https://github.com/KyuDan1/Talk-to-Your-Slides/stargazers)

<div align="center">

**Talk to Your Slides: Language-Driven Agents for Efficient Slide Editing**  
🗒️ Our **research paper** will be published soon!  
🖥️ Our **code** is now available!

</div>

---

## 📖 Overview

Editing presentation slides remains one of the most common and time-consuming tasks faced by millions of users daily, despite significant advances in automated slide generation.

While GUI-based agents have demonstrated visual control capabilities, they often suffer from high computational cost and latency. To address this, we propose **Talk-to-Your-Slides**, an LLM-powered agent that edits slides in active PowerPoint sessions by leveraging structured object-level information—bypassing the need for visual pixel interaction.

Our system introduces a hierarchical editing design, separating high-level semantic planning from low-level object manipulation. This allows:

- 🚀 **34.02% faster** execution  
- 🎯 **34.76% better instruction adherence**  
- 💸 **87.42% cheaper operations**

To evaluate slide editing performance, we present **TSBench**, a human-annotated benchmark with 379 diverse instructions spanning four major categories.


---

## 📚 TSBench Benchmark Dataset

📎 [Download TSBench on Google Drive](https://drive.google.com/drive/folders/1hSjBTCJXiC_rhLGIhLBMqDQpotTr9wiT?usp=sharing)

---

## 🎬 Demo Videos

<div align="center">

[![CamelCase Demo](https://img.youtube.com/vi/9nJ0-yofr7Y/0.jpg)](https://youtu.be/9nJ0-yofr7Y)  
**CamelCase**  
*Prompt:* “Please update all English on ppt slides number 7 to camelCase formatting.”

[![Only English → Blue](https://img.youtube.com/vi/eVSs6xi-bEs/0.jpg)](https://youtu.be/eVSs6xi-bEs)  
**Only English → Blue**  
*Prompt:* “Please change only English into blue color in slide number 3.”

[![Typo Checking Demo](https://img.youtube.com/vi/rBIBsnWX3W0/0.jpg)](https://youtu.be/rBIBsnWX3W0)  
**Typo Checking & Correction**  
*Prompt:* “Please check ppt slides number 4 for any typos or errors, correct them.”

[![Translate to English](https://img.youtube.com/vi/GLS_9xh2C-4/0.jpg)](https://youtu.be/GLS_9xh2C-4)  
**Translate to English**  
*Prompt:* “Please translate ppt slides number 5 into English.”

[![Slide‑Notes Script](https://img.youtube.com/vi/5vzYd5ov_Cs/0.jpg)](https://youtu.be/5vzYd5ov_Cs)  
**Slide Notes Script**  
*Prompt:* “Please create a full script for ppt slides number 3 and add the script to the slide notes.”

</div>

---

## 🛠️ Installation Guide

### 🖥️ Recommended: Python on Windows

⚠️ To allow Python to control PowerPoint via COM interface, you must enable VBA access:

- Open PowerPoint

-  Go to File > Options > Trust Center > Trust Center Settings

- In Macro Settings, make sure to check:
- ✅ "Trust access to the VBA project object model"


1. **Install dependencies:**

```bash
pip install -r requirements.txt
````

2. **Create `credentials.yml`** in the root directory:

```yaml
gpt-4.1-mini:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"

gpt-4.1-nano:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"

gemini-1.5-flash:
  api_key: "YOUR_GEMINI_API_KEY"
```

3. **Create `.env`** in the `pptagent/` directory:

```
# Example .env content
OPENAI_API_KEY=your_key_here
```

4. **Run the system:**

```bash
python pptagent/main.py
```

---
