from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier
import json
import anthropic
import os
import logging
import time
from dotenv import load_dotenv

load_dotenv()
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
GEMINI_API_KEY = os.environ.get('GEMINI')
logging.getLogger('test_Applier').setLevel(logging.DEBUG)

def main(user_input, rule_base_apply:bool = False, log_queue=None, stop_event=None, retry=3):
    import sys

    # Open log file in write mode
    log_file = open(f"./log/output{user_input.replace(' ', '_')}.log", "w", encoding="utf-8")
    # Redirect standard output to log_file
    sys.stdout = log_file

    print("This message will be recorded in output.log")
    
    # Start measurement: Planner
    planner_start_time = time.time()
    planner = Planner()
    plan_json:json = planner(user_input, model_name="gemini-1.5-flash")
    planner_end_time = time.time()
    print("=====PLAN====")
    print(plan_json)

    # Start measurement: Parser
    parser_start_time = time.time()
    parser = Parser(plan_json)
    parsed_json:json = parser.process()
    parser_end_time = time.time()
    print("=====PARSED====")
    print(parsed_json)

    # Start measurement: Processor
    processor_start_time = time.time()
    processor = Processor(parsed_json, model_name = 'gemini-2.5-flash-preview-04-17', api_key=OPENAI_API_KEY)
    processed_json:json = processor.process()
    processor_end_time = time.time()
    print("=====PROCESSED====")
    print(processed_json)   

    # Start measurement: Applier (or test_Applier)
    applier_start_time = time.time()
    if rule_base_apply:
        applier = Applier()
    else:
        applier = test_Applier(model="gemini-2.5-flash-preview-04-17", api_key=OPENAI_API_KEY, retry = retry)
    
    result = applier(processed_json)
    applier_end_time = time.time()

    # End time of total execution
    end_time = time.time()

    # Print time measurement results
    print("\n=====TIME MEASUREMENTS=====")
    print(f"Planner Time:   {planner_end_time - planner_start_time:.4f} seconds")
    print(f"Parser Time:    {parser_end_time - parser_start_time:.4f} seconds")
    print(f"Processor Time: {processor_end_time - processor_start_time:.4f} seconds")
    print(f"Applier Time:   {applier_end_time - applier_start_time:.4f} seconds")
    print(f"Total Time:     {end_time - planner_start_time:.4f} seconds")

    # Close the file after code execution
    log_file.close()


# 테스트 코드
# with open('pptagent/instructions.json', 'r', encoding='utf-8') as f:
#     instructions = json.load(f)
#     print(instructions)
#     for _, inst in instructions.items():
#         try:
#             main(user_input=inst, rule_base_apply=False)
#         except Exception as e:
#             print(f"Error while processing instruction '{inst}': {e}")
#             continue  # 에러가 나면 다음 루프로 넘어감

#main(user_input="Please create a full script for ppt slides number 3 and add the script to the slide notes.", rule_base_apply=False, retry=3)
for i in range(2,50):
    try:
        main(user_input=f"Please translate in English slide number {i}", rule_base_apply=False, retry=4)
    except Exception as e:
        print(f"Error while processing instruction: {e}")
        continue  # Continue to next iteration if error occurs
