from flask import Flask, request, jsonify, render_template
import sys
import os
import importlib.util

app = Flask(__name__)

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
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
            }
            h1 {
                color: #333;
                text-align: center;
            }
            textarea {
                width: 100%;
                height: 150px;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 16px;
                resize: vertical;
            }
            button {
                background-color: #4169E1;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 20px;
                font-size: 16px;
                cursor: pointer;
                margin-top: 10px;
            }
            button:hover {
                background-color: #3154b3;
            }
            #result {
                margin-top: 20px;
                padding: 15px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f9f9f9;
                white-space: pre-wrap;
                display: none;
            }
        </style>
    </head>
    <body>
        <h1>🤖 PPT Agent</h1>
        
        <textarea id="userInput" placeholder="Enter your request here..."></textarea>
        <br>
        <button id="processBtn">Process Request</button>
        
        <div id="result"></div>

        <script>
            document.getElementById('processBtn').addEventListener('click', function() {
                const userInput = document.getElementById('userInput').value.trim();
                
                if (!userInput) {
                    alert('Please enter a request');
                    return;
                }
                
                // Show loading state
                const resultDiv = document.getElementById('result');
                resultDiv.style.display = 'block';
                resultDiv.innerHTML = 'Processing...';
                
                fetch('/run_pipeline', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ user_input: userInput })
                })
                .then(response => response.json())
                .then(data => {
                    resultDiv.innerHTML = data.result || 'Processing complete!';
                })
                .catch(error => {
                    resultDiv.innerHTML = 'Error: ' + error.message;
                    console.error('Error:', error);
                });
            });
        </script>
    </body>
    </html>
    '''

@app.route('/run_pipeline', methods=['POST'])
def run_pipeline():
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
            # main 함수를 직접 호출하고 사용자 입력을 인수로 전달합니다
            result_output = main_module.main(user_input)
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
        return jsonify({
            'status': 'error',
            'result': '',
            'error': str(e)
        }), 500

if __name__ == '__main__':
    try:
        # Use a different port (8080) instead of the default 5000
        # and bind to localhost only for security
        app.run(host='127.0.0.1', port=8080, debug=True)
    except Exception as e:
        print(f"Failed to start the server: {e}")