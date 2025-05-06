import os
import json
import shutil
import win32com.client
import time
import pythoncom
import traceback
from win32com.client import constants

from llm_api import llm_request_with_retries
from classes import Parser
from main_original import main as main_system

# --- Baseline1: Instruction-to-Code 구현 ---
BASELINE_PROMPT = (
    "Create a Python code that can edit PowerPoint presentations by executing the following command:\n"
)
FORMAT_PROMPT = (
    "IMPORTANT: Your response MUST include ONLY valid Python code wrapped in triple backticks with the 'python' language tag. Adhere strictly to this format WITHOUT any additional text or explanation:\n"
    "```python\n"
    "# Include necessary imports, e.g., import win32com.client, from win32com.client import constants\n"
    "# Define functions or classes to interact with PowerPoint via pywin32 (win32com)\n"
    "# Implement any translation logic yourself within the code; do NOT call external translation APIs\n"
    "# Show COM Dispatch, Open, manipulation, Save, and Quit calls as needed\n"
    "```"
)
PARSED_PROMPT = "The following is information parsed from a PPT slide."

def parsing_python(text):
    import re
    python_pattern = r"```python\s*([\s\S]*?)\s*```"
    matches = re.findall(python_pattern, text)
    if matches:
        return matches[0].strip()
    generic = re.findall(r"```\s*([\s\S]*?)\s*```", text)
    if generic:
        return generic[0].strip()
    if "def " in text or "import " in text:
        lines = text.split('\n')
        code_lines = [l for l in lines if l.strip().startswith('#') or re.match(r'^[\w\s\(\)\[\]{}=+\-*/.<>,:\'\"]+$', l)]
        return '\n'.join(code_lines).strip()
    return ""

def baseline1(model_name="gemini-1.5-flash", user_instruction=None):
    parsed_data = Parser(baseline=True)
    prompt = PARSED_PROMPT + str(parsed_data) + BASELINE_PROMPT + user_instruction + FORMAT_PROMPT
    response = llm_request_with_retries(model_name=model_name, request=prompt)
    code = parsing_python(response)
    try:
        exec(code, globals(), locals())
    except Exception as e:
        print(f"Baseline execution error: {e}")
    return code

def main_evaluator(user_input: str):
    try:
        return main_system(user_input=user_input, rule_base_apply=False, retry=4)
    except Exception as e:
        print(f"Main system execution error: {e}")
        raise

# --- PPT 열기/닫기 헬퍼 함수들 ---
def open_presentation(ppt_path: str, max_retries=3):
    for attempt in range(max_retries):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = True
            abs_path = os.path.abspath(ppt_path)
            print(f"Opening presentation: {abs_path}")
            if not os.path.exists(abs_path):
                raise FileNotFoundError(f"PowerPoint file not found: {abs_path}")
            time.sleep(2)
            pres = app.Presentations.Open(abs_path, ReadOnly=False, WithWindow=True)
            time.sleep(1)
            return app, pres
        except Exception as e:
            print(f"Error opening presentation (attempt {attempt+1}/{max_retries}): {e}")
            traceback.print_exc()
            if attempt < max_retries - 1:
                print("Retrying after cleanup...")
                try:
                    kill_powerpoint_instances()
                    time.sleep(3)
                except:
                    pass
            else:
                raise

def save_and_close(app, pres, output_path: str):
    """
    복사해 온 파일(work_ppt) 경로 그대로 저장(pres.Save())하고
    바로 닫고(app.Quit()) COM을 해제합니다.
    """
    try:
        print(f"Saving presentation: {output_path}")
        pres.Save()
        print("Successfully saved.")
    except Exception as e:
        print(f"Save failed: {e}")
    try:
        print("Closing presentation...")
        pres.Close()
        print("Quitting PowerPoint...")
        app.Quit()
    except Exception as e:
        print(f"Error during close/quit: {e}")
        kill_powerpoint_instances()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def kill_powerpoint_instances():
    import subprocess, gc
    try:
        subprocess.run(["taskkill", "/f", "/im", "POWERPNT.EXE"],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        print("Terminated PowerPoint processes.")
        time.sleep(2)
    except Exception as e:
        print(f"Error terminating PowerPoint processes: {e}")
    gc.collect()
    time.sleep(1)

# --- 평가 함수: 입력 폴더의 모든 PPT 파일 동적 처리 ---
def evaluate_on_system(system_func, system_name: str, instructions_file: str,
                       input_ppt_folder: str, output_folder: str):
    with open(instructions_file, 'r', encoding='utf-8') as f:
        instruction_data = json.load(f)
    id_to_prompt = {
        item['id']: item['prompt']
        for cat in instruction_data
        for item in instruction_data[cat]
    }

    ppt_files = sorted(f for f in os.listdir(input_ppt_folder) if f.lower().endswith('.pptx'))
    slide_mapping = {idx: fname for idx, fname in enumerate(ppt_files)}

    os.makedirs(output_folder, exist_ok=True)
    total_runs = 0
    success_runs = 0
    error_log = []

    for slide_num, filename in slide_mapping.items():
        input_ppt = os.path.join(input_ppt_folder, filename)
        pid = slide_num
        raw_prompt = id_to_prompt.get(pid, '')
        if not raw_prompt:
            error_log.append(f"Prompt ID {pid} not found.")
            continue

        prompt = raw_prompt.replace('{slide_num}', str(slide_num))
        print(f"\n{'='*50}")
        print(f"[{system_name}] Slide {slide_num}, Prompt ID {pid}")
        print(f"  -> Prompt: {prompt}")
        print(f"{'='*50}")

        work_ppt = os.path.join(output_folder, f"{system_name}_s{slide_num}_p{pid}.pptx")
        try:
            shutil.copy(input_ppt, work_ppt)
        except Exception as e:
            error_log.append(f"Copy failed: {e}")
            continue

        try:
            app, pres = open_presentation(work_ppt)
            system_func(prompt)
            save_and_close(app, pres, work_ppt)
            success_runs += 1
            print(f"SUCCESS: Slide {slide_num}")
        except Exception as e:
            error_log.append(f"Error on slide {slide_num}: {e}")
            traceback.print_exc()
            kill_powerpoint_instances()

        total_runs += 1
        time.sleep(3)

    if error_log:
        with open(os.path.join(output_folder, f"{system_name}_errors.log"), 'w', encoding='utf-8') as logf:
            logf.write("\n".join(error_log))
        print(f"Errors logged.")

    print(f"\n{'='*60}")
    print(f"{system_name} evaluation: {success_runs}/{total_runs} successful")
    print(f"{'='*60}")
    return success_runs, total_runs

if __name__ == "__main__":
    INSTR_FILE = '../evaluation/mockup_instructions.json'
    PPT_FOLDER = '../evaluation/mockup_ppts'
    OUTPUT_FOLDER = '../results/baseline1'

    try:
        kill_powerpoint_instances()
        import gc; gc.collect()
        time.sleep(2)

        success, total = evaluate_on_system(
            system_func=lambda inp: baseline1(user_instruction=inp),
            system_name='baseline1',
            instructions_file=INSTR_FILE,
            input_ppt_folder=PPT_FOLDER,
            output_folder=OUTPUT_FOLDER
        )
        print(f"Final result: {success}/{total}")
    except KeyboardInterrupt:
        kill_powerpoint_instances()
        print("Interrupted by user.")
    except Exception as e:
        print(f"Unhandled exception: {e}")
        traceback.print_exc()
        kill_powerpoint_instances()
