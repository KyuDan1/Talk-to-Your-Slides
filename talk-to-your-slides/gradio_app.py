import gradio as gr
import os
import json
import logging
import time
import traceback
import pythoncom
from dotenv import load_dotenv
from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier

# Load environment variables
load_dotenv()
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')

# Configure logging
logging.getLogger('test_Applier').setLevel(logging.DEBUG)

def run_pipeline(user_input, retry=3):
    """Modified version of main() that returns the output instead of writing to a file"""
    # Initialize COM library for Windows
    pythoncom.CoInitialize()
    
    # Ensure log directory exists
    os.makedirs("./log", exist_ok=True)
    
    # Create a string buffer to capture output
    output_buffer = []
    
    def append_output(msg):
        output_buffer.append(str(msg))
    
    try:
        # --- 측정 시작: Planner ---
        append_output("Starting processing pipeline...\n")
        planner_start_time = time.time()
        planner = Planner()
        plan_json = planner(user_input, model_name="gemini-1.5-flash")
        planner_end_time = time.time()
        append_output("=====PLAN====")
        append_output(plan_json)

        # --- 측정 시작: Parser ---
        parser_start_time = time.time()
        parser = Parser(plan_json)
        parsed_json = parser.process()
        parser_end_time = time.time()
        append_output("=====PARSED====")
        append_output(parsed_json)

        # --- 측정 시작: Processor ---
        processor_start_time = time.time()
        processor = Processor(parsed_json, model_name='gpt-4.1-mini', api_key=OPENAI_API_KEY)
        processed_json = processor.process()
        processor_end_time = time.time()
        append_output("=====PROCESSED====")
        append_output(processed_json)   

        # --- 측정 시작: Applier (or test_Applier) ---
        applier_start_time = time.time()
        #applier = test_Applier(model="gpt-4.1", api_key=OPENAI_API_KEY, retry=retry)
        applier = test_Applier(model="claude-3.7-sonnet", api_key=ANTHROPIC_API_KEY, retry=retry)
        result = applier(processed_json)
        applier_end_time = time.time()

        # --- 측정 시작: Reporter ---
        reporter_start_time = time.time()
        reporter = Reporter()
        summary = reporter(processed_json, result)
        reporter_end_time = time.time()
        append_output("=====SUMMARY=====")
        append_output(summary)

        # 메모리에 기록
        memory = SharedLogMemory()
        memory = memory(user_input, plan_json, processed_json, result)

        # 전체 실행 종료 시각
        end_time = time.time()

        # --- 시간 측정 결과 출력 ---
        append_output("\n=====TIME MEASUREMENTS=====")
        append_output(f"Planner Time:   {planner_end_time - planner_start_time:.4f} seconds")
        append_output(f"Parser Time:    {parser_end_time - parser_start_time:.4f} seconds")
        append_output(f"Processor Time: {processor_end_time - processor_start_time:.4f} seconds")
        append_output(f"Applier Time:   {applier_end_time - applier_start_time:.4f} seconds")
        append_output(f"Reporter Time:  {reporter_end_time - reporter_start_time:.4f} seconds")
        append_output(f"Total Time:     {end_time - planner_start_time:.4f} seconds")
    
    except Exception as e:
        append_output(f"\nError occurred: {str(e)}\n")
        append_output(traceback.format_exc())
    finally:
        # Always uninitialize COM when done
        pythoncom.CoUninitialize()
    
    # Write to log file as the original function did
    log_file_path = f"./log/output{user_input.replace(' ', '_')}.log"
    with open(log_file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_buffer))
    
    # Return the consolidated output
    return "\n".join(output_buffer)

# Create the Gradio interface
with gr.Blocks(title="Talk-to-Your-Slides") as demo:
    gr.Markdown("# Talk-to-Your-Slides")
    gr.Markdown("Enter your instruction below and click 'Process' to run the Talk-to-Your-Slides.")
    
    with gr.Row():
        with gr.Column(scale=2):
            user_input = gr.Textbox(
                label="User Input", 
                placeholder="Example: Please create a full script for ppt slides number 3 and add the script to the slide notes.",
                lines=3
            )
            
            with gr.Row():
                #rule_base_apply = gr.Checkbox(label="Use Rule-Based Applier", value=False)
                retry = gr.Slider(minimum=1, maximum=5, value=3, step=1, label="Retry Count")
            
            submit_btn = gr.Button("Process", variant="primary")
        
        with gr.Column(scale=3):
            output = gr.Textbox(label="Results", lines=25)
    
    submit_btn.click(
        fn=run_pipeline,
        inputs=[user_input, retry],
        outputs=output
    )

# Launch the app
if __name__ == "__main__":
    demo.launch()