import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import json
import win32com.client
import requests
from PIL import Image, ImageTk
import io
from openai import OpenAI

class PowerPointTextAgent:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint 텍스트 명령 에이전트")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # 스타일 설정
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 12))
        self.style.configure("TLabel", font=("Arial", 12))
        
        # 변수 초기화
        self.api_key = ""
        self.openai_client = None
        self.ppt_app = None
        self.presentation = None
        
        # 전역 네임스페이스 생성 (동적 코드 실행용)
        self.exec_globals = {
            'win32com': win32com,
            'time': time,
            'os': os,
            'print': print
        }
        
        # UI 생성
        self.create_ui()
        
        # PowerPoint 연결 시도
        self.connect_to_powerpoint()

    def create_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # API 키 프레임
        api_frame = ttk.LabelFrame(main_frame, text="OpenAI API 설정", padding=10)
        api_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(api_frame, text="API Key:").pack(side=tk.LEFT, padx=5)
        self.api_key_var = tk.StringVar()
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=40, show="*")
        self.api_key_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(api_frame, text="연결", command=self.connect_api).pack(side=tk.LEFT, padx=5)
        
        # PowerPoint 연결 상태 프레임
        ppt_frame = ttk.Frame(main_frame)
        ppt_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(ppt_frame, text="PowerPoint 연결 상태:").pack(side=tk.LEFT, padx=5)
        self.ppt_status_var = tk.StringVar(value="연결 안됨")
        ttk.Label(ppt_frame, textvariable=self.ppt_status_var).pack(side=tk.LEFT, padx=5)
        ttk.Button(ppt_frame, text="새로고침", command=self.connect_to_powerpoint).pack(side=tk.RIGHT, padx=5)
        
        # 명령 입력 프레임
        input_frame = ttk.LabelFrame(main_frame, text="명령 입력", padding=10)
        input_frame.pack(fill=tk.X, pady=10)
        
        self.command_var = tk.StringVar()
        command_entry = ttk.Entry(input_frame, textvariable=self.command_var, width=50)
        command_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        command_entry.bind("<Return>", lambda event: self.process_command())
        
        ttk.Button(input_frame, text="실행", command=self.process_command).pack(side=tk.RIGHT, padx=5)
        
        # 실행 결과 프레임
        result_frame = ttk.LabelFrame(main_frame, text="실행 결과", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.result_text = tk.Text(result_frame, height=10, width=50)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(self.result_text, orient="vertical", command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # 하단 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="종료", command=self.quit_app).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="기록 지우기", command=lambda: self.result_text.delete(1.0, tk.END)).pack(side=tk.RIGHT, padx=5)
        
        # 상태 표시줄
        self.status_bar = ttk.Label(self.root, text="준비됨", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def connect_api(self):
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("오류", "API 키를 입력해주세요.")
            return
            
        try:
            self.api_key = api_key
            self.openai_client = OpenAI(api_key=api_key)
            # 간단한 테스트 요청으로 API 키 유효성 검증
            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": "테스트"}],
                max_tokens=5
            )
            messagebox.showinfo("성공", "OpenAI API에 성공적으로 연결되었습니다.")
            self.update_status("OpenAI API 연결됨")
        except Exception as e:
            messagebox.showerror("API 연결 오류", f"OpenAI API 연결에 실패했습니다: {str(e)}")
            self.openai_client = None
            self.update_status("API 연결 실패")

    def connect_to_powerpoint(self):
        try:
            # 기존 연결이 있으면 재사용
            if self.ppt_app is None:
                self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            
            # 활성 프레젠테이션 가져오기 시도
            try:
                self.presentation = self.ppt_app.ActivePresentation
                presentation_name = os.path.basename(self.presentation.FullName)
                self.ppt_status_var.set(f"연결됨: {presentation_name}")
                self.update_status(f"PowerPoint에 연결됨: {presentation_name}")
                
                # 전역 네임스페이스에 PowerPoint 객체 추가
                self.exec_globals['ppt_app'] = self.ppt_app
                self.exec_globals['presentation'] = self.presentation
                return True
                
            except:
                self.presentation = None
                self.ppt_status_var.set("연결 안됨: 활성 프레젠테이션 없음")
                self.update_status("오류: 활성 PowerPoint 프레젠테이션이 없습니다")
                return False
                
        except Exception as e:
            self.ppt_app = None
            self.presentation = None
            self.ppt_status_var.set("연결 실패")
            self.update_status(f"PowerPoint 연결 오류: {str(e)}")
            return False

    def process_command(self):
        command_text = self.command_var.get().strip()
        if not command_text:
            self.update_status("명령을 입력해주세요.")
            return
            
        # 먼저 PowerPoint 연결 확인
        if not self.connect_to_powerpoint():
            messagebox.showwarning("경고", "PowerPoint 연결 실패. 활성 프레젠테이션이 있는지 확인하세요.")
            return
            
        # API 연결 확인
        if not self.openai_client:
            messagebox.showwarning("경고", "먼저 OpenAI API에 연결해주세요.")
            return
            
        # 명령 처리 (코드 생성 및 실행)
        self.generate_and_execute_code(command_text)
        
        # 명령 입력 필드 비우기
        self.command_var.set("")

    def generate_and_execute_code(self, command_text):
        self.update_status(f"명령 처리 중: {command_text}")
        self.log_result(f">> 명령: {command_text}")
        
        try:
            # LLM 프롬프트 구성
            system_prompt = """
            당신은 PowerPoint 자동화 전문가입니다. 사용자의 텍스트 명령을 분석하여 Python 코드로 변환해주세요.
            
            코드는 win32com.client를 사용하여 PowerPoint를 조작해야 합니다.
            다음 변수들이 이미 정의되어 있습니다:
            - ppt_app: PowerPoint.Application 객체
            - presentation: 현재 활성화된 PowerPoint 프레젠테이션
            
            코드 작성 시 주의사항:
            1. 코드 블록이나 주석 없이 순수 Python 코드만 제공하세요.
            2. 가능한 모든 오류 상황을 try-except로 처리하세요.
            3. 실행 결과를 출력하려면 print() 함수를 사용하세요.
            4. 결과를 반환하지 말고 직접 작업을 수행하는 코드를 작성하세요.
            5. import 구문을 포함하지 마세요(필요한 모듈은 이미 임포트되어 있습니다).
            """
            
            # API 호출
            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"PowerPoint 명령: {command_text}"}
                ],
                temperature=0.2,
                max_tokens=500
            )
            
            # 응답 처리
            generated_code = response.choices[0].message.content
            
            # 코드 정리 (마크다운 코드 블록 제거)
            if generated_code.startswith("```python"):
                generated_code = generated_code[10:]
            elif generated_code.startswith("```"):
                generated_code = generated_code[3:]
                
            if generated_code.endswith("```"):
                generated_code = generated_code[:-3]
                
            generated_code = generated_code.strip()
            
            # 코드 실행
            self.execute_code(generated_code)
            
        except Exception as e:
            error_message = f"명령 처리 중 오류 발생: {str(e)}"
            self.log_result(f"오류: {error_message}")
            self.update_status(error_message)
            
    def execute_code(self, code):
        # 출력 리디렉션을 위한 설정
        original_print = self.exec_globals['print']
        self.exec_globals['print'] = lambda *args, **kwargs: self.log_result(" ".join(map(str, args)))
        
        try:
            # 코드 실행
            exec(code, self.exec_globals)
            self.update_status("명령 실행 완료")
            
        except Exception as e:
            error_message = f"코드 실행 중 오류 발생: {str(e)}"
            self.log_result(f"오류: {error_message}")
            self.update_status(error_message)
            
        finally:
            # 표준 출력 복구
            self.exec_globals['print'] = original_print

    def log_result(self, text):
        self.result_text.insert(tk.END, text + "\n")
        self.result_text.see(tk.END)
        
    def update_status(self, text):
        self.status_bar.config(text=text)
        print(text)  # 콘솔에도 상태 출력
        
    def quit_app(self):
        # PowerPoint는 닫지 않고 연결만 해제
        self.ppt_app = None
        self.presentation = None
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerPointTextAgent(root)
    root.mainloop()