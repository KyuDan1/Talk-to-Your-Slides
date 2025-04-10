from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
import json

def main(user_input):
    planner = Planner()
    

    
    # 계획 세우기
    plan_json = planner(user_input, model_name="gemini-1.5-flash")
    # planner에서는 json 형식으로 output 할 거임.


    parser = Parser(plan_json)
    parsed_json = parser.process()
    # print(parsed_json)
    # import sys
    # sys.exit()

    processor = Processor(parsed_json)
    processed_json = processor.process()
    print(processed_json)
    import sys
    sys.exit()

    applier = Applier()
    reporter = Reporter()

    memory = SharedLogMemory()
    # print(plan)
    # import sys
    # sys.exit()
    # ppt에서 데이터 파싱하기
    to_process = parser(plan)
    
    # 데이터 processing
    to_apply = processor(to_process)
    
    # ppt에 적용하고, True return하기
    result:bool = applier(to_apply)

    # 이전 내용 모두 넣어서 간추리고 사용자에게 report할 내용 return
    to_report = memory(plan, to_process, to_apply, result)
    
    # 사용자에게 report하기
    reporter(to_report)

#test
main(user_input="please translate ppt slides number 2 in English.")