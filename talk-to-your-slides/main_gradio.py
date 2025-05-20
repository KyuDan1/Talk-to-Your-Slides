from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier
import json
import anthropic
import os
import re
import time
import threading
import logging
import queue
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify
from utils import create_thinking_queue, extract_last_text_content, extract_content_after_edit
load_dotenv()
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')

# 플라스크 앱 생성
app = Flask(__name__)

# 디버깅 로그 활성화
app.logger.setLevel(logging.INFO)

# 생각하는 과정을 저장할 전역 큐
thinking_queue = queue.Queue()
thinking_complete = threading.Event()

# 재시도 카운터 및 최대 재시도 횟수
MAX_RETRIES = 3
# 각 단계별 최대 실행 시간(초)
MAX_STEP_TIME = 300  # 5분

def process_task(user_input, rule_base_apply=False, retry_count=0):
    # 디버깅: 서버 콘솔에 입력값 출력
    app.logger.info(f"프로세스 시작 - 사용자 입력: '{user_input}', rule_base: {rule_base_apply}, 재시도: {retry_count}")
    
    # 재시도 횟수가 최대를 초과하면 중단
    if retry_count >= MAX_RETRIES:
        app.logger.error(f"최대 재시도 횟수({MAX_RETRIES})를 초과했습니다.")
        thinking_queue.put({
            "step": "error",
            "status": "error",
            "message": f"최대 재시도 횟수({MAX_RETRIES})를 초과했습니다. 프로세스를 중단합니다."
        })
        thinking_complete.set()
        return
    
    # COM 초기화 - 스레드에서 PowerPoint 접근을 위해 필요
    try:
        import pythoncom
        pythoncom.CoInitialize()
        app.logger.info("COM 초기화 성공")
    except ImportError:
        app.logger.warning("pythoncom 모듈을 가져올 수 없습니다. COM 초기화 생략")
    
    # 진행 상태와 결과를 저장할 변수들
    result_data = {
        "plan": None,
        "parsed": None,
        "processed": None,
        "result": None,
        "summary": None,
        "times": {}
    }
    
    try:
        # 이전 큐 비우기 (재시도 시)
        if retry_count > 0:
            while not thinking_queue.empty():
                thinking_queue.get()
            thinking_queue.put({
                "step": "restart",
                "status": "info",
                "message": f"오류가 발생하여 프로세스를 재시작합니다. (시도 {retry_count}/{MAX_RETRIES})"
            })
            
        # 타임아웃 설정을 위한 변수
        process_start_time = time.time()
        
        # --- 측정 시작: Planner ---
        thinking_queue.put({
            "step": "planner",
            "status": "thinking",
            "message": "계획 수립 중..."
        })
        
        planner_start_time = time.time()
        planner = Planner()
        app.logger.info(f"Planner 실행 - 사용자 입력: '{user_input}'")
        plan_json = planner(user_input, model_name="gemini-1.5-flash")
        planner_end_time = time.time()
        
        app.logger.info(f"Planner 완료 - 결과: {plan_json[:100]}..." if isinstance(plan_json, str) else f"Planner 완료 - 결과: {str(plan_json)[:100]}...")
        
        result_data["plan"] = plan_json
        result_data["times"]["planner"] = planner_end_time - planner_start_time
        
        printing_text = create_thinking_queue(plan_json)
        thinking_queue.put({
            "step": "planner",
            "status": "complete",
            "message": "계획 수립 완료",
            "data": printing_text,
            "time": planner_end_time - planner_start_time
        })
        
        # --- 측정 시작: Parser ---
        thinking_queue.put({
            "step": "parser",
            "status": "thinking",
            "message": "계획 분석 중..."
        })
        
        parser_start_time = time.time()
        parser = Parser(plan_json)
        parsed_json = parser.process()
        parser_end_time = time.time()
        
        app.logger.info(f"Parser 완료 - 결과: {str(parsed_json)[:100]}...")
        
        result_data["parsed"] = parsed_json
        result_data["times"]["parser"] = parser_end_time - parser_start_time
        
        thinking_queue.put({
            "step": "parser",
            "status": "complete",
            "message": "계획 분석 완료",
            "data": extract_last_text_content(parsed_json),
            "time": parser_end_time - parser_start_time
        })
        
        # --- 측정 시작: Processor ---
        thinking_queue.put({
            "step": "processor",
            "status": "thinking",
            "message": "처리 중..."
        })
        
        processor_start_time = time.time()
        processor = Processor(parsed_json, model_name='gpt-4.1', api_key=OPENAI_API_KEY)
        processed_json = processor.process()
        processor_end_time = time.time()
        
        # 처리 결과 확인
        if processed_json is None:
            app.logger.error("Processor 결과가 None입니다.")
            raise Exception("Processor 결과가 없습니다. 재시작합니다.")
            
        # 타임아웃 체크
        if processor_end_time - processor_start_time > MAX_STEP_TIME:
            app.logger.warning(f"Processor 실행 시간이 {MAX_STEP_TIME}초를 초과했습니다.")
            # 경고는 하지만 일단 계속 진행
        
        app.logger.info(f"Processor 완료")
        
        result_data["processed"] = processed_json
        result_data["times"]["processor"] = processor_end_time - processor_start_time
        
        thinking_queue.put({
            "step": "processor",
            "status": "complete",
            "message": "처리 완료",
            "data": "\n".join(extract_content_after_edit(processed_json)),
            "time": processor_end_time - processor_start_time
        })
        
        # --- 측정 시작: Applier (or test_Applier) ---
        thinking_queue.put({
            "step": "applier",
            "status": "thinking",
            "message": "적용 중..."
        })
        
        applier_start_time = time.time()
        if rule_base_apply:
            applier = Applier()
        else:
            #applier = test_Applier(model="gpt-4.1", api_key=OPENAI_API_KEY)
            applier = test_Applier(model="claude-3.7-sonnet", api_key=ANTHROPIC_API_KEY)
            
        # Applier 실행 및 결과 확인
        result = applier(processed_json)
        applier_end_time = time.time()
        
        # # 'N/A' 관련 오류 검사
        # if isinstance(result, str) and "N/A" in result:
        #     app.logger.error("• manual_review 작업을 'N/A'에 적용합니다.")
        #     raise Exception("Applier에서 'N/A'에 작업을 적용하려는 시도가 있었습니다. 재시작합니다.")
        
        # # 결과가 None이거나 비어있는 경우
        # if result is None or (isinstance(result, (list, dict)) and len(result) == 0):
        #     app.logger.error("Applier 결과가 비어있습니다.")
        #     raise Exception("Applier 결과가 비어있거나 없습니다. 재시작합니다.")
            
        app.logger.info(f"Applier 완료")
        
        result_data["result"] = result
        result_data["times"]["applier"] = applier_end_time - applier_start_time
        
        thinking_queue.put({
            "step": "applier",
            "status": "complete",
            "message": "적용 완료",
            "data": "complete!",
            "time": applier_end_time - applier_start_time
        })
        
        # --- 측정 시작: Reporter ---
        thinking_queue.put({
            "step": "reporter",
            "status": "thinking",
            "message": "보고서 작성 중..."
        })
        
        # 전체 프로세스 타임아웃 체크
        current_time = time.time()
        if current_time - process_start_time > MAX_STEP_TIME * 4:  # 전체 프로세스에 더 긴 시간 허용
            app.logger.error(f"전체 프로세스 실행 시간이 너무 깁니다: {current_time - process_start_time}초")
            raise Exception("프로세스 실행 시간이 너무 깁니다. 재시작합니다.")
        
        reporter_start_time = time.time()
        reporter = Reporter()
        summary = reporter(processed_json, result)
        reporter_end_time = time.time()
        
        # 결과 확인
        if not summary or summary == "N/A" or (isinstance(summary, str) and "manual_review" in summary.lower()):
            app.logger.error("Reporter 결과가 유효하지 않습니다.")
            raise Exception("Reporter 결과가 유효하지 않습니다. 재시작합니다.")
        
        app.logger.info(f"Reporter 완료")
        
        result_data["summary"] = summary
        result_data["times"]["reporter"] = reporter_end_time - reporter_start_time
        
        thinking_queue.put({
            "step": "reporter",
            "status": "complete",
            "message": "보고서 작성 완료",
            "data": summary,
            "time": reporter_end_time - reporter_start_time
        })
        
        # 메모리에 기록
        memory = SharedLogMemory()
        memory = memory(user_input, plan_json, processed_json, result)
        
        # 전체 실행 종료 시각
        end_time = time.time()
        result_data["times"]["total"] = end_time - planner_start_time
        
        # 처리 완료 신호
        thinking_queue.put({
            "step": "complete",
            "status": "complete",
            "message": "모든 처리가 완료되었습니다",
            "data": result_data
        })
        
    except Exception as e:
        app.logger.error(f"오류 발생: {str(e)}")
        # 오류 발생 시 사용자에게 알림
        thinking_queue.put({
            "step": "error",
            "status": "error",
            "message": f"처리 중 오류가 발생했습니다: {str(e)}. 프로세스를 재시작합니다."
        })
        
        # COM 리소스 해제
        try:
            import pythoncom
            pythoncom.CoUninitialize()
            app.logger.info("COM 리소스 해제 완료")
        except ImportError:
            pass
        
        # 재시도 카운터 증가 후 프로세스 재시작
        app.logger.info(f"프로세스 재시작 (시도 {retry_count+1}/{MAX_RETRIES})")
        
        # 잠시 대기 후 재시작 (시스템 리소스가 정리될 시간 제공)
        time.sleep(2)
        
        # 새 스레드에서 프로세스 재시작
        threading.Thread(target=process_task, args=(user_input, rule_base_apply, retry_count + 1)).start()
        return  # 재시도 시작했으므로 현재 함수는 종료
    
    finally:
        # 정상 종료 시에만 COM 리소스 해제 (예외 발생 시에는 이미 해제됨)
        if 'e' not in locals():
            try:
                import pythoncom
                pythoncom.CoUninitialize()
                app.logger.info("COM 리소스 해제 완료")
            except ImportError:
                pass
            
            # 처리 완료 표시
            thinking_complete.set()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    # 요청 내용 로깅
    app.logger.info(f"요청 받음 - 폼 데이터: {request.form}")
    
    user_input = request.form.get('user_input')
    rule_base = request.form.get('rule_base') == 'true'
    
    # 사용자 입력 확인
    if not user_input:
        app.logger.error("사용자 입력이 비어 있습니다.")
        return jsonify({"status": "error", "message": "사용자 입력이 비어 있습니다."})
    
    # 이전 이벤트 초기화
    thinking_complete.clear()
    
    # 큐 비우기 (이전 실행 결과가 있다면)
    while not thinking_queue.empty():
        thinking_queue.get()
    
    # 새 쓰레드에서 처리 시작
    threading.Thread(target=process_task, args=(user_input, rule_base, 0)).start()
    
    return jsonify({"status": "processing"})

@app.route('/thinking_updates')
def thinking_updates():
    if not thinking_queue.empty():
        update = thinking_queue.get()
        app.logger.info(f"업데이트 전송: {update['step']} - {update['status']}")
        return jsonify(update)
    elif thinking_complete.is_set():
        return jsonify({"status": "finished"})
    else:
        return jsonify({"status": "waiting"})

if __name__ == '__main__':
    app.run(debug=True, port=8080)  # 5000 대신 8080 포트 사용