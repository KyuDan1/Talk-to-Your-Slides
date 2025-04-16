document.addEventListener('DOMContentLoaded', function() {
    // 엘리먼트 참조
    const userInputForm = document.getElementById('userInputForm');
    const userInput = document.getElementById('userInput');
    const ruleBase = document.getElementById('ruleBase');
    
    // 뷰 전환 버튼
    const singleViewBtn = document.getElementById('singleViewBtn');
    const expandViewBtn = document.getElementById('expandViewBtn');
    
    // 단일 보드 뷰와 전체 펼치기 뷰
    const singleBoardView = document.getElementById('singleBoardView');
    const expandedView = document.getElementById('expandedView');
    
    // 현재 단계 정보
    const currentIcon = document.getElementById('current-icon');
    const currentTitle = document.getElementById('current-title');
    const currentTime = document.getElementById('current-time');
    const currentThinking = document.getElementById('current-thinking');
    const currentOutput = document.getElementById('current-output');
    
    // 단계별 아이콘과 타이틀 매핑
    const stepInfo = {
        'planner': { icon: '🧠', title: '계획 수립' },
        'parser': { icon: '📊', title: '계획 분석' },
        'processor': { icon: '⚙️', title: '처리' },
        'applier': { icon: '🔄', title: '적용' },
        'reporter': { icon: '📝', title: '보고서 작성' },
        'complete': { icon: '✅', title: '최종 결과' }
    };
    
    // 현재 활성화된 단계
    let currentStep = '';
    // 단계별 결과 데이터 저장
    const stepResults = {};
    
    // 모든 단계 배열
    const steps = ['planner', 'parser', 'processor', 'applier', 'reporter', 'complete'];
    
    // 초기화: 모든 단계 비활성화 및 진행 상태 초기화
    function initializeUI() {
        // 단일 보드 뷰에서 초기 상태 설정
        currentIcon.textContent = '⏳';
        currentTitle.textContent = '대기 중...';
        currentTime.textContent = '';
        currentThinking.style.display = 'none';
        currentOutput.textContent = '';
        currentOutput.classList.remove('visible');
        
        // 진행 상태 초기화
        document.querySelectorAll('.progress-step').forEach(step => {
            step.classList.remove('active', 'complete');
        });
        
        // 전체 펼치기 뷰 초기화
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
        
        // 결과 데이터 초기화
        Object.keys(stepResults).forEach(key => delete stepResults[key]);
        currentStep = '';
    }
    
    // 뷰 전환 버튼 이벤트 처리
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
    
    // 폼 제출 이벤트 처리
    userInputForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        if (!userInput.value.trim()) {
            alert('지시사항을 입력해주세요.');
            return;
        }
        
        console.log('폼 제출 - 사용자 입력:', userInput.value);
        
        // UI 초기화
        initializeUI();
        
        // 기본 뷰로 전환
        singleViewBtn.click();
        
        // 폼 데이터 준비
        const formData = new FormData();
        formData.append('user_input', userInput.value);
        formData.append('rule_base', ruleBase.checked);
        
        // 처리 시작 표시
        currentIcon.textContent = '⏳';
        currentTitle.textContent = '요청 처리 중...';
        currentThinking.style.display = 'flex';
        
        // 서버에 요청 전송
        fetch('/process', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            console.log('서버 응답:', data);
            
            if (data.status === 'processing') {
                // 폴링 시작
                pollThinkingUpdates();
            } else {
                alert('오류: ' + JSON.stringify(data));
            }
        })
        .catch(error => {
            console.error('오류:', error);
            alert('처리 중 오류가 발생했습니다.');
        });
    });
    
    // 생각 과정 업데이트 폴링
    function pollThinkingUpdates() {
        fetch('/thinking_updates')
            .then(response => response.json())
            .then(data => {
                console.log('업데이트:', data);
                
                if (data.status === 'waiting') {
                    // 대기 중, 계속 확인
                    setTimeout(pollThinkingUpdates, 500);
                    return;
                }
                
                if (data.status === 'error') {
                    // 오류 발생
                    console.error('오류:', data.message);
                    alert('오류: ' + data.message);
                    return;
                }
                
                if (data.status === 'finished') {
                    // 모든 처리 완료
                    console.log('모든 처리 완료');
                    return;
                }
                
                // 데이터 처리
                handleThinkingUpdate(data);
                
                // 계속 폴링
                setTimeout(pollThinkingUpdates, 500);
            })
            .catch(error => {
                console.error('폴링 오류:', error);
                setTimeout(pollThinkingUpdates, 1000);
            });
    }
    
    // 단계 업데이트 처리
    function handleThinkingUpdate(data) {
        const step = data.step;
        const status = data.status;
        
        // 새로운 단계가 시작되면 이전 단계 완료 처리
        if (currentStep && currentStep !== step && status === 'thinking') {
            // 진행 상태 표시에서 이전 단계 완료 표시
            const prevStepElement = document.querySelector(`.progress-step[data-step="${currentStep}"]`);
            if (prevStepElement) {
                prevStepElement.classList.remove('active');
                prevStepElement.classList.add('complete');
            }
        }
        
        // 최종 완료 단계
        if (step === 'complete') {
            // 단일 보드 뷰에서 마지막 단계 표시
            updateCurrentStepDisplay('complete', 'complete', data);
            
            // 진행 상태 표시에서 모든 단계 완료 표시
            document.querySelectorAll('.progress-step').forEach(el => {
                el.classList.remove('active');
                el.classList.add('complete');
            });
            
            // 전체 펼치기 뷰 업데이트
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
                    timeElement.textContent = `총 소요 시간: ${data.data.times.total.toFixed(2)}초`;
                }
            }
            
            // 모든 단계 활성화 (전체 펼치기 뷰용)
            steps.forEach(s => {
                if (s !== 'complete' && stepResults[s]) {
                    const stepEl = document.getElementById(s);
                    if (stepEl) stepEl.classList.add('active');
                }
            });
            
            return;
        }
        
        // 새로운 단계 시작 또는 현재 단계 업데이트
        if (status === 'thinking') {
            // 단계가 변경됨
            if (currentStep !== step) {
                currentStep = step;
                
                // 단일 보드 뷰 업데이트
                updateCurrentStepDisplay(step, 'thinking');
                
                // 진행 상태 표시 업데이트
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
            
            // 전체 펼치기 뷰 업데이트
            const stepElement = document.getElementById(step);
            if (stepElement) {
                stepElement.classList.add('active');
                
                const thinkingElement = document.getElementById(`${step}-thinking`);
                if (thinkingElement) {
                    thinkingElement.style.display = 'flex';
                }
            }
        } else if (status === 'complete') {
            // 완료된 단계의 데이터 저장
            stepResults[step] = {
                data: data.data,
                time: data.time
            };
            
            // 단일 보드 뷰 업데이트
            updateCurrentStepDisplay(step, 'complete', data);
            
            // 전체 펼치기 뷰 업데이트
            const stepElement = document.getElementById(step);
            if (stepElement) {
                const thinkingElement = document.getElementById(`${step}-thinking`);
                if (thinkingElement) {
                    thinkingElement.style.display = 'none';
                }
                
                const outputElement = document.getElementById(`${step}-output`);
                if (outputElement) {
                    // JSON 데이터 포맷팅
                    let formattedData = typeof data.data === 'object' 
                        ? JSON.stringify(data.data, null, 2) 
                        : data.data;
                    
                    outputElement.textContent = formattedData;
                    outputElement.classList.add('visible');
                }
                
                const timeElement = document.getElementById(`${step}-time`);
                if (timeElement && data.time) {
                    timeElement.textContent = `소요 시간: ${data.time.toFixed(2)}초`;
                }
            }
        }
    }
    
    // 단일 보드 뷰의 현재 단계 표시 업데이트
    function updateCurrentStepDisplay(step, status, data) {
        // 페이드인 효과를 위한 opacity 변경
        currentTitle.classList.add('changing');
        
        setTimeout(() => {
            // 아이콘 및 제목 업데이트
            if (stepInfo[step]) {
                currentIcon.textContent = stepInfo[step].icon;
                currentTitle.textContent = stepInfo[step].title;
            }
            
            // 상태에 따른 표시 변경
            if (status === 'thinking') {
                currentThinking.style.display = 'flex';
                currentOutput.classList.remove('visible');
                currentTime.textContent = '';
            } else if (status === 'complete') {
                currentThinking.style.display = 'none';
                
                if (data && data.data) {
                    // JSON 데이터 포맷팅
                    let formattedData = typeof data.data === 'object' 
                        ? JSON.stringify(data.data, null, 2) 
                        : data.data;
                    
                    currentOutput.textContent = formattedData;
                    currentOutput.classList.add('visible');
                }
                
                if (data && data.time) {
                    currentTime.textContent = `소요 시간: ${data.time.toFixed(2)}초`;
                }
            }
            
            // 페이드인 효과 완료
            currentTitle.classList.remove('changing');
        }, 300); // 페이드 아웃 후 내용 변경
    }
    
    // 초기 UI 상태 설정
    initializeUI();
    singleViewBtn.click(); // 기본 뷰로 단일 보드 뷰 선택
});