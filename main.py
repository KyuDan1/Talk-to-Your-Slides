import os
import time
import random
import json
import numpy as np
import cv2
import pyautogui
import gymnasium as gym
from gymnasium import spaces
import torch
import base64
from io import BytesIO
from PIL import Image, ImageGrab
import requests
import stable_baselines3 as sb3
from stable_baselines3 import PPO
from stable_baselines3.common.vec_env import DummyVecEnv
from transformers import AutoModelForCausalLM, AutoTokenizer
import pygetwindow as gw

# API 키 설정 (실제 사용 시 환경 변수로 관리하는 것을 추천)
LLM_API_KEY = "your_llm_api_key"  # 예: OpenAI API 키
VLM_API_KEY = "your_vlm_api_key"  # 예: OpenAI API 키 또는 다른 Vision API 키

class PowerPointAction:
    """PowerPoint에서 수행할 수 있는 액션 정의"""
    def __init__(self, action_type, params=None):
        self.action_type = action_type  # click, type, scroll, drag 등
        self.params = params or {}      # x, y 좌표, 텍스트 등
    
    def execute(self):
        """액션 실행"""
        if self.action_type == "click":
            pyautogui.click(self.params.get("x"), self.params.get("y"))
            return True
        
        elif self.action_type == "right_click":
            pyautogui.rightClick(self.params.get("x"), self.params.get("y"))
            return True
        
        elif self.action_type == "double_click":
            pyautogui.doubleClick(self.params.get("x"), self.params.get("y"))
            return True
        
        elif self.action_type == "type":
            pyautogui.typewrite(self.params.get("text"))
            return True
        
        elif self.action_type == "hotkey":
            pyautogui.hotkey(*self.params.get("keys"))
            return True
        
        elif self.action_type == "scroll":
            pyautogui.scroll(self.params.get("amount"))
            return True
        
        elif self.action_type == "drag":
            pyautogui.moveTo(self.params.get("start_x"), self.params.get("start_y"))
            pyautogui.dragTo(self.params.get("end_x"), self.params.get("end_y"), 
                             duration=self.params.get("duration", 0.5))
            return True
            
        return False


class ScreenCapture:
    """화면 캡처 및 이미지 처리"""
    @staticmethod
    def find_powerpoint_window():
        """PowerPoint 창을 찾아 반환"""
        # PowerPoint 창 이름에 따라 검색 (다양한 버전 지원)
        powerpoint_titles = ["PowerPoint", "Microsoft PowerPoint", "프레젠테이션", "Presentation"]
        
        for title in powerpoint_titles:
            windows = gw.getWindowsWithTitle(title)
            for window in windows:
                if window.visible and not window.isMinimized:
                    return window
        
        # PowerPoint 창을 찾지 못한 경우
        return None
    
    @staticmethod
    def capture_screen(powerpoint_only=True):
        """현재 화면을 캡처하여 PIL 이미지로 반환
        
        Args:
            powerpoint_only (bool): True면 PowerPoint 창만 캡처, False면 전체 화면 캡처
        
        Returns:
            PIL.Image: 캡처된 이미지
        """
        if powerpoint_only:
            # PowerPoint 창 찾기
            ppt_window = ScreenCapture.find_powerpoint_window()
            
            if ppt_window:
                # 창 위치와 크기 가져오기
                left, top, right, bottom = ppt_window.left, ppt_window.top, ppt_window.right, ppt_window.bottom
                
                # 창 영역만 캡처
                screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
                return screenshot
            else:
                print("경고: PowerPoint 창을 찾을 수 없습니다. 전체 화면을 캡처합니다.")
        
        # PowerPoint 창을 찾지 못했거나 전체 화면 캡처 모드일 경우
        screenshot = ImageGrab.grab()
        return screenshot
    
    @staticmethod
    def image_to_base64(image):
        """PIL 이미지를 base64 인코딩 문자열로 변환"""
        buffered = BytesIO()
        image.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode("utf-8")


class LLMModule:
    """자연어 명령을 이해하고 액션 계획을 생성하는 LLM 모듈"""
    def __init__(self, api_key=LLM_API_KEY, model_name="gpt-4"):
        self.api_key = api_key
        self.model_name = model_name
        
    def _create_prompt(self, user_command, screen_description=None, manual_context=None):
        """LLM 프롬프트 생성"""
        prompt = f"""
        당신은 PowerPoint를 제어하는 AI 어시스턴트입니다. 
        사용자의 명령을 이해하고 PowerPoint GUI에서 수행할 정확한 액션을 결정해야 합니다.
        
        사용 가능한 액션 유형:
        - click: GUI 요소 클릭 (x, y 좌표 필요)
        - right_click: 우클릭 (x, y 좌표 필요)
        - double_click: 더블클릭 (x, y 좌표 필요)
        - type: 텍스트 입력 (text 필요)
        - hotkey: 단축키 입력 (keys 목록 필요, 예: ["ctrl", "s"])
        - scroll: 스크롤 (amount 필요, 양수는 위로, 음수는 아래로)
        - drag: 드래그 (start_x, start_y, end_x, end_y, duration 필요)
        
        사용자 명령: {user_command}
        """
        
        if screen_description:
            prompt += f"\n현재 화면 상태: {screen_description}\n"
            
        if manual_context:
            prompt += f"\nPowerPoint 매뉴얼 관련 정보: {manual_context}\n"
            
        prompt += """
        JSON 형식으로 다음 단계의 액션을 반환하세요:
        {
            "action_type": "액션 유형",
            "params": {
                // 액션 유형에 따른 필요 파라미터
            },
            "reasoning": "이 액션을 선택한 이유"
        }
        """
        
        return prompt
        
    def get_action_plan(self, user_command, screen_description=None, manual_context=None):
        """사용자 명령에 따른 액션 계획 생성"""
        prompt = self._create_prompt(user_command, screen_description, manual_context)
        
        # OpenAI API 호출 (다른 LLM API로 대체 가능)
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        data = {
            "model": self.model_name,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2
        }
        
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=data
        )
        
        if response.status_code == 200:
            response_json = response.json()
            action_json = response_json["choices"][0]["message"]["content"]
            try:
                action_plan = json.loads(action_json)
                return action_plan
            except json.JSONDecodeError:
                return {"error": "LLM 응답을 JSON으로 파싱할 수 없습니다."}
        else:
            return {"error": f"LLM API 오류: {response.status_code}"}


class VLMModule:
    """화면 이미지를 분석하는 VLM 모듈"""
    def __init__(self, api_key=VLM_API_KEY, model_name="gpt-4-vision-preview"):
        self.api_key = api_key
        self.model_name = model_name
        
    def analyze_screen(self, screenshot):
        """화면 이미지 분석"""
        base64_image = ScreenCapture.image_to_base64(screenshot)
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        data = {
            "model": self.model_name,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": """현재 PowerPoint 화면을 분석하고 다음 정보를 제공해주세요:
                            1. 현재 보이는 주요 UI 요소들 (메뉴, 버튼, 슬라이드 등)의 위치와 상태
                            2. 화면에서 보이는 텍스트 내용
                            3. 현재 PowerPoint의 상태 (편집 모드, 프레젠테이션 모드 등)
                            JSON 형식으로 응답해주세요."""
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            "temperature": 0.2
        }
        
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=data
        )
        
        if response.status_code == 200:
            response_json = response.json()
            analysis_text = response_json["choices"][0]["message"]["content"]
            try:
                # VLM이 JSON으로 응답하지 않을 수도 있으므로 예외 처리
                analysis = json.loads(analysis_text)
                return analysis
            except json.JSONDecodeError:
                # 텍스트 분석 그대로 반환
                return {"description": analysis_text}
        else:
            return {"error": f"VLM API 오류: {response.status_code}"}


class ManualKnowledgeBase:
    """PowerPoint 매뉴얼 기반 지식베이스"""
    def __init__(self, manual_file_path="powerpoint_manual.json"):
        self.manual_data = {}
        try:
            with open(manual_file_path, 'r', encoding='utf-8') as f:
                self.manual_data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            print(f"경고: {manual_file_path} 파일을 로드할 수 없습니다. 빈 지식베이스로 시작합니다.")
            
    def get_relevant_info(self, command):
        """명령어와 관련된 매뉴얼 정보 검색"""
        # 실제 구현에서는 임베딩 기반 검색이나 키워드 매칭 등 사용
        # 간단한 예시 구현:
        relevant_sections = []
        
        # 명령어에서 키워드 추출
        keywords = command.lower().split()
        
        for section in self.manual_data.get("sections", []):
            for keyword in keywords:
                if (keyword in section.get("title", "").lower() or 
                    keyword in section.get("content", "").lower()):
                    relevant_sections.append(section)
                    break
                    
        return relevant_sections


class PowerPointEnv(gym.Env):
    """PowerPoint 강화학습 환경"""
    def __init__(self, llm_module, vlm_module, manual_kb):
        super(PowerPointEnv, self).__init__()
        
        self.llm_module = llm_module
        self.vlm_module = vlm_module
        self.manual_kb = manual_kb
        
        # 액션 공간 정의 (간소화된 버전)
        # 실제로는 더 복잡한 액션 공간이 필요할 수 있음
        self.action_space = spaces.Dict({
            "action_type": spaces.Discrete(7),  # 7가지 액션 타입
            "params": spaces.Dict({
                "x": spaces.Box(low=0, high=1920, shape=(1,), dtype=np.int32),
                "y": spaces.Box(low=0, high=1080, shape=(1,), dtype=np.int32),
                "text": spaces.Text(10),  # 최대 10자 텍스트
                "amount": spaces.Box(low=-10, high=10, shape=(1,), dtype=np.int32)
            })
        })
        
        # 관측 공간 정의 (화면 이미지)
        self.observation_space = spaces.Box(
            low=0, high=255, shape=(1080, 1920, 3), dtype=np.uint8
        )
        
        self.current_command = None
        self.episode_steps = 0
        self.max_episode_steps = 10
        
    def reset(self, seed=None, options=None):
        """환경 초기화"""
        super().reset(seed=seed)
        
        # PowerPoint 초기 상태로 (새 프레젠테이션)
        # 실제로는 PowerPoint 실행 및 초기화 코드가 필요
        
        # 훈련용 명령어 선택 (실제로는 명령어 데이터셋 필요)
        commands = [
            "새 슬라이드 추가",
            "슬라이드에 제목 추가",
            "글꼴 크기 변경",
            "이미지 삽입",
            "슬라이드 쇼 시작"
        ]
        self.current_command = random.choice(commands)
        
        # 현재 화면 캡처
        screenshot = ScreenCapture.capture_screen()
        screen_array = np.array(screenshot)
        
        self.episode_steps = 0
        
        return screen_array, {"command": self.current_command}
        
    def step(self, action_dict):
        """환경에서 한 스텝 진행"""
        self.episode_steps += 1
        
        # 딕셔너리에서 액션 파라미터 추출
        action_type_idx = action_dict["action_type"]
        params = action_dict["params"]
        
        # 인덱스를 액션 타입 문자열로 변환
        action_types = ["click", "right_click", "double_click", "type", "hotkey", "scroll", "drag"]
        action_type = action_types[action_type_idx]
        
        # PowerPointAction 객체 생성 및 실행
        action = PowerPointAction(action_type, params)
        success = action.execute()
        
        # 잠시 대기 (UI 업데이트 대기)
        time.sleep(0.5)
        
        # 새 화면 상태 관측
        screenshot = ScreenCapture.capture_screen()
        screen_array = np.array(screenshot)
        
        # VLM으로 화면 분석
        screen_analysis = self.vlm_module.analyze_screen(screenshot)
        
        # 보상 계산
        reward = self._calculate_reward(success, screen_analysis)
        
        # 에피소드 종료 여부 확인
        done = self.episode_steps >= self.max_episode_steps
        
        return screen_array, reward, done, False, {"command": self.current_command, "analysis": screen_analysis}
    
    def _calculate_reward(self, action_success, screen_analysis):
        """보상 계산"""
        if not action_success:
            return -1.0  # 액션 실패 패널티
        
        # 화면 분석 기반 보상
        reward = 0.0
        
        # 매뉴얼 지식베이스에서 관련 정보 가져오기
        relevant_info = self.manual_kb.get_relevant_info(self.current_command)
        
        # 명령어와 화면 상태 비교하여 목표 달성 여부 확인
        # (실제로는 더 복잡한 로직이 필요)
        if "error" in screen_analysis:
            reward -= 0.5
        else:
            # 예: "새 슬라이드 추가" 명령에 대해 화면에 새 슬라이드가 표시되는지 확인
            description = screen_analysis.get("description", "")
            if "새 슬라이드" in self.current_command and "새 슬라이드" in description:
                reward += 2.0
            elif "제목 추가" in self.current_command and "텍스트 상자" in description:
                reward += 2.0
            # 기타 명령어별 보상 로직...
            
            # 기본 보상
            reward += 0.1
        
        return reward


class PPOAgent:
    """PPO 강화학습 에이전트"""
    def __init__(self, env, model_path=None):
        self.env = env
        
        # 모델 설정
        if model_path and os.path.exists(model_path):
            self.model = PPO.load(model_path, env=env)
            print(f"모델을 {model_path}에서 로드했습니다.")
        else:
            self.model = PPO(
                "MultiInputPolicy",
                env,
                verbose=1,
                learning_rate=0.0003,
                n_steps=2048,
                batch_size=64,
                n_epochs=10,
                gamma=0.99,
                tensorboard_log="./ppt_agent_tensorboard/"
            )
            print("새 모델을 초기화했습니다.")
    
    def train(self, total_timesteps=10000):
        """에이전트 훈련"""
        print(f"{total_timesteps} 타임스텝 동안 훈련을 시작합니다...")
        self.model.learn(total_timesteps=total_timesteps)
        print("훈련 완료!")
        
    def save(self, model_path="models/ppt_agent_model"):
        """모델 저장"""
        os.makedirs(os.path.dirname(model_path), exist_ok=True)
        self.model.save(model_path)
        print(f"모델을 {model_path}에 저장했습니다.")
        
    def predict(self, observation, command):
        """주어진 관측과 명령어에 대한 액션 예측"""
        # 관측과 명령어 결합
        action, _ = self.model.predict(observation)
        return action


class PowerPointAgent:
    """파워포인트 GUI 자동화 에이전트 통합 클래스"""
    def __init__(self, use_trained_model=False, model_path="models/ppt_agent_model"):
        # 구성 요소 초기화
        self.llm_module = LLMModule()
        self.vlm_module = VLMModule()
        self.manual_kb = ManualKnowledgeBase()
        
        # 강화학습 환경 설정
        self.env = DummyVecEnv([lambda: PowerPointEnv(
            self.llm_module, self.vlm_module, self.manual_kb
        )])
        
        # 에이전트 초기화
        if use_trained_model and os.path.exists(model_path):
            self.agent = PPOAgent(self.env, model_path)
        else:
            self.agent = PPOAgent(self.env)
            
    def train(self, total_timesteps=10000):
        """에이전트 훈련"""
        self.agent.train(total_timesteps)
        self.agent.save()
        
    def execute_command(self, command):
        """자연어 명령 실행"""
        print(f"명령어 실행: {command}")
        
        # 현재 화면 캡처
        screenshot = ScreenCapture.capture_screen()
        
        # VLM으로 화면 분석
        screen_analysis = self.vlm_module.analyze_screen(screenshot)
        screen_description = screen_analysis.get("description", "")
        
        # 매뉴얼에서 관련 정보 검색
        manual_info = self.manual_kb.get_relevant_info(command)
        manual_context = "\n".join([f"{info.get('title')}: {info.get('content')}" 
                                   for info in manual_info])
        
        # LLM으로 액션 계획 생성
        action_plan = self.llm_module.get_action_plan(
            command, screen_description, manual_context
        )
        
        if "error" in action_plan:
            print(f"오류: {action_plan['error']}")
            return False
            
        # 액션 실행
        action = PowerPointAction(
            action_plan["action_type"], 
            action_plan.get("params", {})
        )
        
        print(f"실행할 액션: {action_plan['action_type']}")
        print(f"액션 파라미터: {action_plan.get('params', {})}")
        print(f"액션 근거: {action_plan.get('reasoning', '')}")
        
        return action.execute()


def main():
    """메인 함수"""
    print("PowerPoint GUI 자동화 에이전트 시작...")
    
    # 에이전트 초기화 (학습된 모델 사용 여부 선택)
    agent = PowerPointAgent(use_trained_model=False)
    
    # 훈련 모드
    train_mode = input("에이전트를 훈련하시겠습니까? (y/n): ").lower() == 'y'
    
    if train_mode:
        timesteps = int(input("훈련할 타임스텝 수 (기본: 10000): ") or "10000")
        agent.train(total_timesteps=timesteps)
        print("훈련 완료!")
    
    # 실행 모드
    while True:
        command = input("\n파워포인트 명령어를 입력하세요 (종료하려면 'exit' 입력): ")
        
        if command.lower() == 'exit':
            break
            
        success = agent.execute_command(command)
        
        if success:
            print("명령어 실행 성공!")
        else:
            print("명령어 실행 실패")


if __name__ == "__main__":
    main()