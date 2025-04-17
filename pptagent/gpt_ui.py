from flask import Flask, render_template_string, request, Response
from classes import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
from test_Applier import test_Applier
import os, time
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')

# Inline HTML template with SSE integration
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PPT Agent UI</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 600px; margin: 40px auto; }
        textarea { width: 100%; height: 80px; }
        #status { margin-top: 20px; padding: 10px; border: 1px solid #ddd; height: 200px; overflow-y: auto; background: #f9f9f9; }
        .running { color: blue; }
        .done { color: green; }
    </style>
</head>
<body>
    <h2>PPT Agent Controller</h2>
    <form id="controlForm">
        <label for="user_input">Enter command:</label><br>
        <textarea id="user_input"></textarea><br>
        <button type="submit">Run</button>
    </form>
    <div id="status"></div>

    <script>
        const form = document.getElementById('controlForm');
        const statusDiv = document.getElementById('status');
        let source;

        form.addEventListener('submit', e => {
            e.preventDefault();
            if (source) source.close();
            statusDiv.innerHTML = '';
            const userInput = encodeURIComponent(document.getElementById('user_input').value);
            source = new EventSource(`/stream?user_input=${userInput}`);

            source.onmessage = function(event) {
                const p = document.createElement('p');
                const msg = event.data;
                if (msg.includes('running')) p.className = 'running';
                else if (msg.includes('done') || msg.includes('Completed')) p.className = 'done';
                p.textContent = msg;
                statusDiv.appendChild(p);
                statusDiv.scrollTop = statusDiv.scrollHeight;
            };
        });
    </script>
</body>
</html>
'''

app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/stream')
def stream():
    user_input = request.args.get('user_input', '')

    def event_stream():
        # Stage: Planner
        yield 'data: Planner running...\n\n'
        planner = Planner()
        plan = planner(user_input, model_name="gemini-1.5-flash")
        time.sleep(0.2)
        yield 'data: Planner done\n\n'

        # Stage: Parser
        yield 'data: Parser running...\n\n'
        parser = Parser(plan)
        parsed = parser.process()
        time.sleep(0.2)
        yield 'data: Parser done\n\n'

        # Stage: Processor
        yield 'data: Processor running...\n\n'
        processor = Processor(parsed, model_name='gpt-4.1-mini', api_key=OPENAI_API_KEY)
        processed = processor.process()
        time.sleep(0.2)
        yield 'data: Processor done\n\n'

        # Stage: Applier
        yield 'data: Applier running...\n\n'
        applier = test_Applier(model="gpt-4.1", api_key=OPENAI_API_KEY, retry=3)
        result = applier(processed)
        time.sleep(0.2)
        yield 'data: Applier done\n\n'

        # Stage: Reporter
        yield 'data: Reporter running...\n\n'
        reporter = Reporter()
        summary = reporter(processed, result)
        time.sleep(0.2)
        yield f'data: Reporter done - {summary}\n\n'

        # Completion
        yield 'data: All stages completed.\n\n'

    return Response(event_stream(), mimetype='text/event-stream')

if __name__ == '__main__':
    app.run(debug=True, port=8080)
