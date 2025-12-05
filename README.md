---

<div align="center">

# 📜 *Talk to Your Slides:*

## **Language-Driven Agents for Efficient Slide Editing**
[![Stars](https://img.shields.io/github/stars/KyuDan1/Talk-to-Your-Slides?style=social)](https://github.com/KyuDan1/Talk-to-Your-Slides/stargazers)


---
📄 **[Research Paper (arXiv preprint)](https://arxiv.org/abs/2505.11604)**

</div>

---

> **Note:** We will release TSBench-Hard version soon!

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

📎 [Download TSBench-Hard on Google Drive]() (To be updated)

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

1. Open PowerPoint
2. Go to **File > Options > Trust Center > Trust Center Settings**
3. In **Macro Settings**, check:
   - ✅ "Trust access to the VBA project object model"

### 📦 Setup Instructions

#### Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

**Note:** If you encounter issues with package installation, install these core packages:
```bash
pip install openai==1.74.0 google-generativeai anthropic python-pptx Flask python-dotenv pyyaml
```

#### Step 2: Configure API Keys

**Option A: Using credentials.yml (Recommended)**

Copy the example credentials file:
```bash
cp credentials.yml.example credentials.yml
```

Edit `credentials.yml` with your API keys:
```yaml
gpt-4.1-mini:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"

gpt-4.1:
  api_key:  "YOUR_OPENAI_API_KEY"
  base_url: "https://api.openai.com/v1"

gemini-1.5-flash:
  api_key: "YOUR_GEMINI_API_KEY"

claude-3.7-sonnet:
  api_key: "YOUR_ANTHROPIC_API_KEY"
```

**Option B: Using .env file**

Create a `.env` file in the `pptagent/` directory:
```bash
cd pptagent
cat > .env << EOF
OPENAI_API_KEY=your_openai_key_here
ANTHROPIC_API_KEY=your_anthropic_key_here
GEMINI_API_KEY=your_gemini_key_here
EOF
```

#### Step 3: Run the System

**Web UI (Flask) - Recommended for interactive use:**
```bash
python pptagent/main_flask.py
```
Then open your browser to `http://localhost:8080`

**CLI Mode - For batch processing:**
```bash
cd pptagent
python main_cli.py
```

**Quick Start (shows usage):**
```bash
python pptagent/main.py
```

### 🔧 Project Structure

```
Talk-to-Your-Slides/
├── pptagent/
│   ├── main.py              # Entry point (shows usage)
│   ├── main_flask.py        # Web UI server (Flask)
│   ├── main_cli.py          # CLI interface
│   ├── classes.py           # Core PPT agent classes
│   ├── test_Applier.py      # Applier implementations
│   ├── llm_api.py           # LLM API wrappers
│   ├── gemini_api.py        # Gemini-specific API
│   ├── utils.py             # Utility functions
│   ├── prompt.py            # System prompts
│   └── templates/           # Flask HTML templates
├── credentials.yml.example  # Example API credentials
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

### 🎯 Supported Models

- **OpenAI**: GPT-4.1, GPT-4.1-mini, GPT-4.1-nano
- **Google**: Gemini 1.5 Flash, Gemini 2.5 Flash
- **Anthropic**: Claude 3.7 Sonnet

### 💡 Usage Examples

**Example 1: Translate slide content**
```
"Translate all text content on slide 1 into Korean."
```

**Example 2: Fix typos**
```
"Check slide 4 for any typos or errors and correct them."
```

**Example 3: Change formatting**
```
"Change all English text to blue color on slide 3."
```

See demo videos below for more examples!

### 🐛 Troubleshooting

**Issue: ModuleNotFoundError for openai or google.generativeai**
```bash
# Solution: Install missing packages
pip install openai==1.74.0 google-generativeai
```

**Issue: FileNotFoundError for credentials.yml**
```bash
# Solution: Create credentials file from example
cp credentials.yml.example credentials.yml
# Then edit credentials.yml with your API keys
```

**Issue: COM error on Windows**
- Make sure PowerPoint is installed
- Enable VBA access (see installation guide above)
- Run Python as Administrator if needed

**Issue: Flask server not starting**
```bash
# Check if port 8080 is available
# Try a different port by editing main_flask.py line 341:
# app.run(debug=True, port=8081)  # Change to different port
```

### 🏗️ Code Architecture

The system follows a hierarchical pipeline:

1. **Planner**: Analyzes user request and creates high-level plan
2. **Parser**: Parses the plan into structured tasks
3. **Processor**: Processes each task with contextual information
4. **Applier**: Applies changes to PowerPoint slides via COM/python-pptx
5. **Reporter**: Generates summary of changes made

Each component is modular and can be extended independently.

---

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
---
