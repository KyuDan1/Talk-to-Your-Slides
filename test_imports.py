# test_imports.py
try:
    import pyautogui
    print("✓ PyAutoGUI 설치됨")
    
    import pygetwindow
    print("✓ PyGetWindow 설치됨")
    
    from PIL import ImageGrab
    print("✓ PIL/Pillow 설치됨")
    
    import numpy as np
    print("✓ NumPy 설치됨")
    
    import cv2
    print("✓ OpenCV 설치됨")
    
    print("\n모든 핵심 패키지가 정상적으로 설치되었습니다!")
except ImportError as e:
    print(f"오류 발생: {e}")