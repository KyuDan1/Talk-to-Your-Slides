from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier, test_json_Applier
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

    # 로그 파일을 쓰기 모드로 엽니다.
    log_file = open(f"./log/output{user_input.replace(' ', '_')}.log", "w", encoding="utf-8")
    # 표준 출력을 log_file로 재지정합니다.
    sys.stdout = log_file

    print("이 메시지는 output.log 파일에 기록됩니다.")
    
    # --- 측정 시작: Planner ---
    planner_start_time = time.time()
    planner = Planner()
    plan_json, planner_input_tokens, planner_output_tokens, planner_price = planner(user_input, model_name="gemini-1.5-flash")
    planner_end_time = time.time()
    print("=====PLAN====")
    print(plan_json)

    # --- 측정 시작: Parser ---
    parser_start_time = time.time()
    parser = Parser(plan_json)
    parsed_json:json = parser.process()
    parser_end_time = time.time()
    print("=====PARSED====")
    print(parsed_json)
    
    # --- 측정 시작: Processor ---
    processor_start_time = time.time()
    processor = Processor(parsed_json, model_name = 'gemini-2.5-flash-preview-04-17', 
                          api_key=GEMINI_API_KEY#OPENAI_API_KEY
                          )
    processed_json, processor_input_tokens, processor_output_tokens, processor_price = processor.process()
    processor_end_time = time.time()
    print("=====PROCESSED====")
    print(processed_json)   
    
    # --- 측정 시작: Applier (or test_Applier) ---
    applier_start_time = time.time()
    if rule_base_apply:
        applier = Applier()
    else:
        applier= test_json_Applier(model='gemini-2.5-flash-preview-04-17' #"gpt-4.1", 
                                    ,api_key=GEMINI_API_KEY
                                    ,retry = retry) #test_Applier(model="gpt-4.1", api_key=OPENAI_API_KEY, retry = retry)
    
    result ,applier_input_tokens, applier_output_tokens, applier_total_cost = applier(processed_json)
    applier_end_time = time.time()
    print(result)

    # --- 측정 시작: Reporter ---
    #reporter_start_time = time.time()
    #reporter = Reporter()
    #summary = reporter(processed_json, result)
    #reporter_end_time = time.time()
    #print("=====SUMMARY=====")
    #print(summary)

    # 메모리에 기록
    #memory = SharedLogMemory()
    #memory = memory(user_input, plan_json, processed_json, result)

    # 전체 실행 종료 시각
    end_time = time.time()

    # --- 시간 측정 결과 출력 ---
    print("\n=====TIME MEASUREMENTS=====")
    print(f"Planner Time:   {planner_end_time - planner_start_time:.4f} seconds")
    print(f"Parser Time:    {parser_end_time - parser_start_time:.4f} seconds")
    print(f"Processor Time: {processor_end_time - processor_start_time:.4f} seconds")
    print(f"Applier Time:   {applier_end_time - applier_start_time:.4f} seconds")
    #print(f"Reporter Time:  {reporter_end_time - reporter_start_time:.4f} seconds")
    print(f"Total Time:     {end_time - planner_start_time:.4f} seconds")

    # 코드 실행 후 파일을 닫습니다.
    log_file.close()

    total_input_token = planner_input_tokens + processor_input_tokens + applier_input_tokens
    total_output_token = planner_output_tokens + processor_output_tokens + applier_output_tokens
    total_price = planner_price + processor_price + applier_total_cost
    
    return result, total_input_token, total_output_token, total_price

# result, total_input_token, total_output_token, total_price = main("Translate all text content on slide 1 into Korean.")
