from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier
import json
import anthropic
import os
import logging
from dotenv import load_dotenv
load_dotenv()
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
logging.getLogger('test_Applier').setLevel(logging.DEBUG)
def main(user_input):
    # 계획짜는 class (외부 LLM 사용)
    # input: 사용자 명령 output: 계획 json
    planner = Planner()
    plan_json:json = planner(user_input, model_name="gemini-1.5-flash")
    # print(plan_json)
    # ppt의 요소 가져오는 애. python 코드를 실행함. 정해진 형식이 있어 LLM 사용 안함.
    parser = Parser(plan_json)
    parsed_json:json = parser.process()

    # ppt에서 가져온 요소를 plan에 맞춰서 행동하는 애.(번역, 요약, 검색 등) (외부 LLM 사용)
    processor = Processor(parsed_json)
    processed_json:json = processor.process()
    # print(processed_json)
    # import sys
    # sys.exit()
    # processed 된 내용을 ppt에 적용하는 애. (ppt에 적용하는 python 코드를 짜서 실행하는 애.)
    # 현재는 rule-base로 python 코드를 짜고 있는데, 이게 오류율이 상당히 높음.
    # 이 부분을 데이터 수집해서 LLM으로 돌리면 더 안정적일 것으로 예상함.
    # 부분으로 쪼개서 각각 적용하는 python 코드를 짜게 하는게 더 안정적일 것으로 예상함.
    
    #applier = Applier()
    
    applier = test_Applier(api_key=ANTHROPIC_API_KEY)
    result = applier(processed_json)
    #print(processed_json)
    # 진행 된 사항을 사용자에게 리포트하는 애. (외부 LLM 사용)
    reporter = Reporter()
    summary = reporter(processed_json, result)
    print(summary)

    # 이전 내용을 모두 저장. context를 가지고 있음.
    memory = SharedLogMemory()
    memory = memory(user_input, plan_json, processed_json, result)
    
#test
main(user_input="please translate ppt slides number 3 in Korean.")
