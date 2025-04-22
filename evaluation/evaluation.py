import sys
sys.path.append('C:/Users/wjdrb/PPTAgent-4.16/pptagent')

from main_original import *

import json
import os
from typing import Callable
def run_evaluation(eval_system):
    """
    Run evaluations on the eval_system for all 37 slides.
    
    Args:
        eval_system: A function that takes a prompt parameter and returns a response
    """
    # Load the instruction data
    with open('./evaluation/test_instructions.json', 'r', encoding='utf-8') as f:
        instruction_data = json.load(f)
    
    # Create a mapping of ID to prompt
    id_to_prompt = {}
    for category in instruction_data:
        for item in instruction_data[category]:
            id_to_prompt[item['id']] = item['prompt']
    
    # Define the slide mapping as provided
    slide_mapping = {
        (1, 2, 3, 4): [0, 23, 26, 14],  # 4개, 영어 아닌 텍스트 있음
        (5, 6): [15, 30],  # 2개, 푸른 계열 도형들 중 하나만 빨간 도형 있음
        (7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18): [2, 3, 4, 6, 13, 16, 19, 29, 31, 32, 34, 35],  # 12개
        (19, 20, 21): [5, 17, 18],  # 3개, typo가 있는 텍스트박스가 있고 좌우에 크기가 다른 이미지가 하나씩 있음
        (22,): [7],  # 1개, 대문자 및 punctuation을 잘못한 text가 있음
        (23, 24, 25, 26): [8, 10, 11, 1],  # 간단한 영어가 있되 32pt가 아니고 line space가 1.5가 아님
        (27, 28, 29): [9, 24, 36],  # 3개, bullet list 텍스트가 있음
        (30, 31): [12, 33],  # 2개, 흰색 바탕에 흰색에 가까운 텍스트가 있음
        (32, 33, 34): [20, 21, 25],  # 3개, 이미지와 텍스트 쌍이 2쌍 있음
        (35,): [22],  # 1개, 서로다른 도형 두개가 있음
        (36,): [27],  # 1개, 간단한 데이터가 있는 표가 있음
        (37,): [28]   # 1개, 간단한 chart가 있음
    }
    
    # Track total evaluations
    total_count = 0
    
    # Run evaluations for each slide
    for slide_range, prompt_ids in slide_mapping.items():
        for slide_num in slide_range:
            for prompt_id in prompt_ids:
                # Get the prompt template and fill in the slide number
                prompt_template = id_to_prompt[prompt_id]
                prompt = prompt_template.replace('{slide_num}', str(slide_num))
                
                # Print current evaluation information
                print(f"Evaluating Slide {slide_num} with Prompt ID {prompt_id}: {prompt[:50]}..." 
                      if len(prompt) > 50 else f"Evaluating Slide {slide_num} with Prompt ID {prompt_id}: {prompt}")
                
                # Execute the system with the prompt
                #eval_system(prompt=prompt)
                eval_system(user_input=prompt, rule_base_apply=False, retry=4)

                total_count += 1
    
    print(f"\nCompleted {total_count} evaluations across 37 slides.")

if __name__ == "__main__":
    # Example usage with a mock evaluation system
    # def mock_eval_system(prompt):
    #     """Mock function to simulate the evaluation system"""
    #     pass  # Just execute, no return value
    
    run_evaluation(main)