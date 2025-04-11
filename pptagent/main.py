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
    #print(processed_json)

    # import sys
    # sys.exit()

    applier = Applier()
    result = applier(processed_json)
    #print(f"결과: {result}")
    import sys
    sys.exit()
    reporter = Reporter()
    summary = reporter(processed_json, result)
    print(summary)
    
    memory = SharedLogMemory()
    # 이전 내용 모두 넣어서 간추리고 사용자에게 report할 내용 return
    memory = memory(user_input, plan_json, processed_json, result)
    
#test
main(user_input="please translate ppt slides number 2 in English.")