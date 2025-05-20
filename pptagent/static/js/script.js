document.addEventListener('DOMContentLoaded', function() {
    // Element references
    const userInputForm = document.getElementById('userInputForm');
    const userInput = document.getElementById('userInput');
    const ruleBase = document.getElementById('ruleBase');
    
    // View toggle buttons
    const singleViewBtn = document.getElementById('singleViewBtn');
    const expandViewBtn = document.getElementById('expandViewBtn');
    
    // Single board view and expanded view
    const singleBoardView = document.getElementById('singleBoardView');
    const expandedView = document.getElementById('expandedView');
    
    // Current step information
    const currentIcon = document.getElementById('current-icon');
    const currentTitle = document.getElementById('current-title');
    const currentTime = document.getElementById('current-time');
    const currentThinking = document.getElementById('current-thinking');
    const currentOutput = document.getElementById('current-output');
    
    // Step icon and title mapping
    const stepInfo = {
        'planner': { icon: '🧠', title: 'Planning' },
        'parser': { icon: '📊', title: 'Plan Analysis' },
        'processor': { icon: '⚙️', title: 'Processing' },
        'applier': { icon: '🔄', title: 'Application' },
        'reporter': { icon: '📝', title: 'Report Writing' },
        'complete': { icon: '✅', title: 'Final Result' }
    };
    
    // Currently active step
    let currentStep = '';
    // Step result data storage
    const stepResults = {};
    
    // All steps array
    const steps = ['planner', 'parser', 'processor', 'applier', 'reporter', 'complete'];
    
    // Initialize UI
    function initializeUI() {
        // Initialize single board view
        currentIcon.textContent = '⏳';
        currentTitle.textContent = 'Waiting...';
        currentTime.textContent = '';
        currentThinking.style.display = 'none';
        currentOutput.textContent = '';
        currentOutput.classList.remove('visible');
        
        // Initialize progress steps
        document.querySelectorAll('.progress-step').forEach(step => {
            step.classList.remove('active', 'complete');
        });
        
        // Initialize expanded view
        steps.forEach(step => {
            const stepElement = document.getElementById(step);
            if (stepElement) {
                stepElement.classList.remove('active');
                
                const outputElement = document.getElementById(`${step}-output`);
                if (outputElement) {
                    outputElement.textContent = '';
                    outputElement.classList.remove('visible');
                }
                
                const thinkingElement = document.getElementById(`${step}-thinking`);
                if (thinkingElement) {
                    thinkingElement.style.display = 'none';
                }
                
                const timeElement = document.getElementById(`${step}-time`);
                if (timeElement) {
                    timeElement.textContent = '';
                }
            }
        });
        
        // Initialize result data
        Object.keys(stepResults).forEach(key => delete stepResults[key]);
        currentStep = '';
    }
    
    // View toggle button event handling
    singleViewBtn.addEventListener('click', function() {
        singleViewBtn.classList.add('active');
        expandViewBtn.classList.remove('active');
        singleBoardView.classList.add('active');
        expandedView.classList.remove('active');
    });
    
    expandViewBtn.addEventListener('click', function() {
        expandViewBtn.classList.add('active');
        singleViewBtn.classList.remove('active');
        expandedView.classList.add('active');
        singleBoardView.classList.remove('active');
    });
    
    // Form submission event handling
    userInputForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        if (!userInput.value.trim()) {
            alert('Please enter instructions.');
            return;
        }
        
        console.log('Form submission - User input:', userInput.value);
        
        // Initialize UI
        initializeUI();
        
        // Switch to default view
        singleViewBtn.click();
        
        // Prepare form data
        const formData = new FormData();
        formData.append('user_input', userInput.value);
        formData.append('rule_base', ruleBase.checked);
        
        // Show processing started
        currentIcon.textContent = '⏳';
        currentTitle.textContent = 'Processing request...';
        currentThinking.style.display = 'flex';
        
        // Send request to server
        fetch('/process', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            console.log('Server response:', data);
            
            if (data.status === 'processing') {
                // Start polling
                pollThinkingUpdates();
            } else {
                alert('Error: ' + JSON.stringify(data));
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('Processing error occurred.');
        });
    });
    
    // Thinking process update polling
    function pollThinkingUpdates() {
        fetch('/thinking_updates')
            .then(response => response.json())
            .then(data => {
                console.log('Update:', data);
                
                if (data.status === 'waiting') {
                    // Waiting, continue checking
                    setTimeout(pollThinkingUpdates, 500);
                    return;
                }
                
                if (data.status === 'error') {
                    // Error occurred
                    console.error('Error:', data.message);
                    alert('Error: ' + data.message);
                    return;
                }
                
                if (data.status === 'finished') {
                    // All processing completed
                    console.log('All processing completed');
                    return;
                }
                
                // Process data
                handleThinkingUpdate(data);
                
                // Continue polling
                setTimeout(pollThinkingUpdates, 500);
            })
            .catch(error => {
                console.error('Polling error:', error);
                setTimeout(pollThinkingUpdates, 1000);
            });
    }
    
    // Step update processing
    function handleThinkingUpdate(data) {
        const step = data.step;
        const status = data.status;
        
        // If a new step starts, complete previous step
        if (currentStep && currentStep !== step && status === 'thinking') {
            // Mark previous step as complete in progress display
            const prevStepElement = document.querySelector(`.progress-step[data-step="${currentStep}"]`);
            if (prevStepElement) {
                prevStepElement.classList.remove('active');
                prevStepElement.classList.add('complete');
            }
        }
        
        // Final complete step
        if (step === 'complete') {
            // Mark last step in single board view
            updateCurrentStepDisplay('complete', 'complete', data);
            
            // Mark all steps as complete in progress display
            document.querySelectorAll('.progress-step').forEach(el => {
                el.classList.remove('active');
                el.classList.add('complete');
            });
            
            // Update expanded view
            const finalElement = document.getElementById('final-result');
            if (finalElement) {
                finalElement.classList.add('active');
                
                const outputElement = document.getElementById('final-output');
                if (outputElement && data.data && data.data.summary) {
                    outputElement.textContent = JSON.stringify(data.data.summary, null, 2);
                    outputElement.classList.add('visible');
                }
                
                const timeElement = document.getElementById('total-time');
                if (timeElement && data.data && data.data.times) {
                    timeElement.textContent = `Total processing time: ${data.data.times.total.toFixed(2)} seconds`;
                }
            }
            
            // Activate all steps (for expanded view)
            steps.forEach(s => {
                if (s !== 'complete' && stepResults[s]) {
                    const stepEl = document.getElementById(s);
                    if (stepEl) stepEl.classList.add('active');
                }
            });
            
            return;
        }
        
        // Start new step or update current step
        if (status === 'thinking') {
            // Step changed
            if (currentStep !== step) {
                currentStep = step;
                
                // Update single board view
                updateCurrentStepDisplay(step, 'thinking');
                
                // Update progress display
                document.querySelectorAll('.progress-step').forEach(el => {
                    if (el.dataset.step === step) {
                        el.classList.add('active');
                    } else if (steps.indexOf(el.dataset.step) < steps.indexOf(step)) {
                        el.classList.remove('active');
                        el.classList.add('complete');
                    } else {
                        el.classList.remove('active', 'complete');
                    }
                });
            }
            
            // Update expanded view
            const stepElement = document.getElementById(step);
            if (stepElement) {
                stepElement.classList.add('active');
                
                const thinkingElement = document.getElementById(`${step}-thinking`);
                if (thinkingElement) {
                    thinkingElement.style.display = 'flex';
                }
            }
        } else if (status === 'complete') {
            // Store data for completed step
            stepResults[step] = {
                data: data.data,
                time: data.time
            };
            
            // Update single board view
            updateCurrentStepDisplay(step, 'complete', data);
            
            // Update expanded view
            const stepElement = document.getElementById(step);
            if (stepElement) {
                const thinkingElement = document.getElementById(`${step}-thinking`);
                if (thinkingElement) {
                    thinkingElement.style.display = 'none';
                }
                
                const outputElement = document.getElementById(`${step}-output`);
                if (outputElement) {
                    // JSON data formatting
                    let formattedData = typeof data.data === 'object' 
                        ? JSON.stringify(data.data, null, 2) 
                        : data.data;
                    
                    outputElement.textContent = formattedData;
                    outputElement.classList.add('visible');
                }
                
                const timeElement = document.getElementById(`${step}-time`);
                if (timeElement && data.time) {
                    timeElement.textContent = `Processing time: ${data.time.toFixed(2)} seconds`;
                }
            }
        }
    }
    
    // Single board view current step display update
    function updateCurrentStepDisplay(step, status, data) {
        // Fade-in effect for opacity change
        currentTitle.classList.add('changing');
        
        setTimeout(() => {
            // Update icon and title
            if (stepInfo[step]) {
                currentIcon.textContent = stepInfo[step].icon;
                currentTitle.textContent = stepInfo[step].title;
            }
            
            // Update display based on status
            if (status === 'thinking') {
                currentThinking.style.display = 'flex';
                currentOutput.classList.remove('visible');
                currentTime.textContent = '';
            } else if (status === 'complete') {
                currentThinking.style.display = 'none';
                
                if (data && data.data) {
                    // JSON data formatting
                    let formattedData = typeof data.data === 'object' 
                        ? JSON.stringify(data.data, null, 2) 
                        : data.data;
                    
                    currentOutput.textContent = formattedData;
                    currentOutput.classList.add('visible');
                }
                
                if (data && data.time) {
                    currentTime.textContent = `Processing time: ${data.time.toFixed(2)} seconds`;
                }
            }
            
            // Fade-in effect completed
            currentTitle.classList.remove('changing');
        }, 300); // Fade-out delay before content change
    }
    
    // Initialize UI state
    initializeUI();
    singleViewBtn.click(); // Select single board view by default
});