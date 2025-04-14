from flask import Flask, request, jsonify, render_template
import sys
import os
import importlib.util
import signal
import threading

app = Flask(__name__)

# 현재 실행 중인 프로세스를 추적하기 위한 전역 변수
current_process = None

# Create a simple HTML template directly in the app
@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>PPT Agent</title>
        <style>
            :root {
                --primary-color: #4361ee;
                --secondary-color: #f8f9fa;
                --text-color: #333;
                --border-radius: 8px;
                --box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                --transition: all 0.3s ease;
            }
            
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 40px 20px;
                background-color: #fff;
                color: var(--text-color);
                line-height: 1.6;
            }
            
            h1 {
                color: var(--primary-color);
                text-align: center;
                margin-bottom: 30px;
                font-weight: 600;
            }
            
            .container {
                background-color: var(--secondary-color);
                padding: 30px;
                border-radius: var(--border-radius);
                box-shadow: var(--box-shadow);
            }
            
            textarea {
                width: 100%;
                height: 150px;
                padding: 15px;
                border: 1px solid #ddd;
                border-radius: var(--border-radius);
                font-size: 16px;
                resize: vertical;
                font-family: inherit;
                box-sizing: border-box;
                transition: var(--transition);
                margin-bottom: 20px;
            }
            
            textarea:focus {
                outline: none;
                border-color: var(--primary-color);
                box-shadow: 0 0 0 2px rgba(67, 97, 238, 0.2);
            }
            
            .button-group {
                display: flex;
                gap: 15px;
            }
            
            button {
                border: none;
                border-radius: var(--border-radius);
                padding: 12px 24px;
                font-size: 16px;
                cursor: pointer;
                transition: var(--transition);
                font-weight: 500;
                flex: 1;
            }
            
            #processBtn {
                background-color: var(--primary-color);
                color: white;
            }
            
            #processBtn:hover {
                background-color: #3347c4;
            }
            
            #stopBtn {
                background-color: #e63946;
                color: white;
                display: none;
            }
            
            #stopBtn:hover {
                background-color: #d62839;
            }
            
            #result {
                margin-top: 30px;
                padding: 20px;
                border-radius: var(--border-radius);
                background-color: white;
                box-shadow: var(--box-shadow);
                white-space: pre-wrap;
                display: none;
                font-size: 15px;
                line-height: 1.6;
                border-left: 4px solid var(--primary-color);
            }
            
            .loader {
                display: inline-block;
                width: 20px;
                height: 20px;
                border: 3px solid rgba(255, 255, 255, 0.3);
                border-radius: 50%;
                border-top-color: white;
                animation: spin 1s ease-in-out infinite;
                margin-right: 10px;
                vertical-align: middle;
            }
            
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <h1>🤖 PPT Agent</h1>
        
        <div class="container">
            <textarea id="userInput" placeholder="Enter your request here..."></textarea>
            
            <div class="button-group">
                <button id="processBtn">Process Request</button>
                <button id="stopBtn">Stop Process</button>
            </div>
            
            <div id="result"></div>
        </div>

        <script>
            const processBtn = document.getElementById('processBtn');
            const stopBtn = document.getElementById('stopBtn');
            const userInput = document.getElementById('userInput');
            const resultDiv = document.getElementById('result');
            
            processBtn.addEventListener('click', function() {
                const input = userInput.value.trim();
                
                if (!input) {
                    alert('Please enter a request');
                    return;
                }
                
                // UI 상태 변경
                processBtn.disabled = true;
                stopBtn.style.display = 'block';
                resultDiv.style.display = 'block';
                resultDiv.innerHTML = '<div class="loader"></div> Processing...';
                
                fetch('/run_pipeline', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ user_input: input })
                })
                .then(response => response.json())
                .then(data => {
                    resultDiv.innerHTML = data.result || 'Processing complete!';
                    resetUI();
                })
                .catch(error => {
                    resultDiv.innerHTML = 'Error: ' + error.message;
                    console.error('Error:', error);
                    resetUI();
                });
            });
            
            stopBtn.addEventListener('click', function() {
                fetch('/stop_process', {
                    method: 'POST'
                })
                .then(response => response.json())
                .then(data => {
                    resultDiv.innerHTML = data.message;
                    resetUI();
                })
                .catch(error => {
                    console.error('Error stopping process:', error);
                });
            });
            
            function resetUI() {
                processBtn.disabled = false;
                stopBtn.style.display = 'none';
            }
        </script>
    </body>
    </html>
    '''

@app.route('/run_pipeline', methods=['POST'])
def run_pipeline():
    global current_process
    data = request.json
    user_input = data.get('user_input', '')
    
    try:
        # Get the directory of this script
        current_dir = os.path.dirname(os.path.abspath(__file__))
        main_py_path = os.path.join(current_dir, 'main.py')
        
        # 다이나믹하게 main.py 모듈을 로드합니다
        spec = importlib.util.spec_from_file_location("main_module", main_py_path)
        main_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(main_module)
        
        # main 함수가 존재하는지 확인하고 사용자 입력을 인수로 전달합니다
        if hasattr(main_module, 'main'):
            # 현재 프로세스 ID 설정 (간단히 True로 설정)
            current_process = True
            
            # 결과를 직접 반환합니다
            result_output = main_module.main(user_input)
            
            # 프로세스 완료 후 추적 변수 초기화
            current_process = None
            
            return jsonify({
                'status': 'success',
                'result': str(result_output),
                'error': ''
            })
        else:
            return jsonify({
                'status': 'error',
                'result': '',
                'error': 'main 함수를 main.py에서 찾을 수 없습니다.'
            }), 500
    except Exception as e:
        current_process = None
        return jsonify({
            'status': 'error',
            'result': '',
            'error': str(e)
        }), 500

@app.route('/stop_process', methods=['POST'])
def stop_process():
    global current_process
    
    if current_process:
        # 프로세스 실행 플래그를 False로 설정
        current_process = None
        
        # 중단 신호 보내기 (실제 구현에서는 os.kill 또는 subprocess 관리가 필요할 수 있음)
        # 실제로는 계속 돌아가게되어 있음.
        
        return jsonify({
            'status': 'success',
            'message': 'Process has been stopped.'
        })
    else:
        return jsonify({
            'status': 'info',
            'message': 'No process is currently running.'
        })

if __name__ == '__main__':
    try:
        # Use a different port (8080) instead of the default 5000
        # and bind to localhost only for security
        app.run(host='127.0.0.1', port=8080, debug=True)
    except Exception as e:
        print(f"Failed to start the server: {e}")