# PPTAgent

### recommend python in Windows.

### conda environment
```bash
pip install -r 'requirements.txt'
```
- Then make 'credentials.yml' on pptagent-4.16/pptagent.
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
- .env file in pptagent
```bash
python pptagent/main.py
```
### Overall
<img src="fig1.png">

