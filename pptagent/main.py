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
logging.getLogger('test_Applier').setLevel(logging.DEBUG)

def main(user_input, rule_base_apply:bool = False):
    import sys

    # 로그 파일을 쓰기 모드로 엽니다.
    log_file = open(f"./log/output{user_input.replace(' ', '_')}.log", "w", encoding="utf-8")
    # 표준 출력을 log_file로 재지정합니다.
    sys.stdout = log_file

    print("이 메시지는 output.log 파일에 기록됩니다.")
    
    # --- 측정 시작: Planner ---
    planner_start_time = time.time()
    planner = Planner()
    plan_json:json = planner(user_input, model_name="gemini-1.5-flash")
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
    processor = Processor(parsed_json, model_name = 'gpt-4.1-mini', api_key=OPENAI_API_KEY)
    processed_json:json = processor.process()
    processor_end_time = time.time()
    print("=====PROCESSED====")
    print(processed_json)   

    # --- 측정 시작: Applier (or test_Applier) ---
    applier_start_time = time.time()
    if rule_base_apply:
        applier = Applier()
    else:
        applier = test_Applier(model="gpt-4.1", api_key=OPENAI_API_KEY)
    result = applier(processed_json)
    applier_end_time = time.time()

    # --- 측정 시작: Reporter ---
    reporter_start_time = time.time()
    reporter = Reporter()
    summary = reporter(processed_json, result)
    reporter_end_time = time.time()
    print("=====SUMMARY=====")
    print(summary)

    # 메모리에 기록
    memory = SharedLogMemory()
    memory = memory(user_input, plan_json, processed_json, result)

    # 전체 실행 종료 시각
    end_time = time.time()

    # --- 시간 측정 결과 출력 ---
    print("\n=====TIME MEASUREMENTS=====")
    print(f"Planner Time:   {planner_end_time - planner_start_time:.4f} seconds")
    print(f"Parser Time:    {parser_end_time - parser_start_time:.4f} seconds")
    print(f"Processor Time: {processor_end_time - processor_start_time:.4f} seconds")
    print(f"Applier Time:   {applier_end_time - applier_start_time:.4f} seconds")
    print(f"Reporter Time:  {reporter_end_time - reporter_start_time:.4f} seconds")
    print(f"Total Time:     {end_time - planner_start_time:.4f} seconds")

    # 코드 실행 후 파일을 닫습니다.
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

main(user_input="Please split ppt slides number 9 into two slides.", rule_base_apply=False)
