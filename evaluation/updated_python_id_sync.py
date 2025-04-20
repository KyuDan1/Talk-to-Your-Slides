import json
import argparse
import os

def update_ids_and_sync(dataset_file, groups_file=None, output_dataset_file=None, output_groups_file=None):
    """
    데이터셋과 그룹화된 ID를 함께 업데이트하는 함수
    
    Args:
        dataset_file (str): 데이터셋 JSON 파일 경로
        groups_file (str, optional): 그룹화된 ID JSON 파일 경로
        output_dataset_file (str, optional): 업데이트된 데이터셋을 저장할 파일 경로
        output_groups_file (str, optional): 업데이트된 그룹을 저장할 파일 경로
        
    Returns:
        dict: 업데이트된 데이터셋, 그룹화된 ID, ID 매핑 정보
    """
    try:
        # 출력 파일이 지정되지 않은 경우 입력 파일로 설정
        if output_dataset_file is None:
            output_dataset_file = dataset_file
            
        # 1. 데이터셋 파일 읽기
        with open(dataset_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 2. 현재 ID 추출 및 정렬
        current_ids = []
        for category in data:
            for item in data[category]:
                current_ids.append(item['id'])
        current_ids.sort()
        
        # 3. ID 매핑 생성 (기존 ID -> 새 ID)
        id_mapping = {id_val: i for i, id_val in enumerate(current_ids)}
        
        # 4. 수정된 그룹화된 데이터 정의
        if groups_file:
            # 그룹 파일에서 데이터 읽기
            with open(groups_file, 'r', encoding='utf-8') as f:
                groups = json.load(f)
        else:
            # 코드에 정의된 그룹 데이터 사용
            groups = {
                'one': [1, 24, 29, 15],
                'two': [16, 34],
                'three': [3, 4, 5, 7, 14, 17, 20, 33, 35, 36, 41, 43],
                'four': [6, 18, 19],
                'five': [8],
                'six': [9, 11, 12, 2],
                'seven': [10, 25, 44],
                'eight': [13, 39],
                'nine': [21, 22, 26],
                'ten': [23],
                'eleven': [30],
                'twelve': [31]
            }
        
        # 5. 데이터셋의 ID 업데이트
        for category in data:
            for item in data[category]:
                item['id'] = id_mapping[item['id']]
        
        # 6. 그룹화된 데이터의 ID 업데이트
        updated_groups = {}
        for group_name, group_ids in groups.items():
            updated_groups[group_name] = [id_mapping[id_val] for id_val in group_ids if id_val in id_mapping]
        
        # 7. 누락된 ID 확인
        all_group_ids = [id_val for group_ids in groups.values() for id_val in group_ids]
        missing_in_dataset = [id_val for id_val in all_group_ids if id_val not in id_mapping]
        
        if missing_in_dataset:
            print(f"경고: 다음 ID들은 그룹에는 있지만 데이터셋에는 없습니다: {missing_in_dataset}")
        
        # 8. 업데이트된 데이터셋 저장
        with open(output_dataset_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        # 9. 업데이트된 그룹 저장
        if output_groups_file:
            with open(output_groups_file, 'w', encoding='utf-8') as f:
                json.dump(updated_groups, f, indent=2, ensure_ascii=False)
        
        # 10. 업데이트된 결과 반환
        return {
            'updated_data': data,
            'updated_groups': updated_groups,
            'id_mapping': id_mapping,
            'missing_ids': missing_in_dataset
        }
        
    except Exception as e:
        print(f"오류 발생: {e}")
        raise

def create_groups_file(output_file):
    """
    수정된 그룹 데이터를 JSON 파일로 저장하는 함수
    
    Args:
        output_file (str): 출력 파일 경로
    """
    groups = {
        'one': [1, 24, 29, 15],
        'two': [16, 34],
        'three': [3, 4, 5, 7, 14, 17, 20, 33, 35, 36, 41, 43],
        'four': [6, 18, 19],
        'five': [8],
        'six': [9, 11, 12, 2],
        'seven': [10, 25, 44],
        'eight': [13, 39],
        'nine': [21, 22, 26],
        'ten': [23],
        'eleven': [30],
        'twelve': [31]
    }
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(groups, f, indent=2, ensure_ascii=False)
    
    print(f"그룹화 파일이 생성되었습니다: {output_file}")

def main():
    parser = argparse.ArgumentParser(description="데이터셋 ID와 그룹화된 ID를 0부터 시작하는 오름차순으로 업데이트합니다.")
    parser.add_argument('--dataset', required=True, help='데이터셋 JSON 파일 경로')
    parser.add_argument('--groups', help='그룹화된 ID JSON 파일 경로')
    parser.add_argument('--create-groups', action='store_true', help='수정된 그룹 데이터로 그룹 파일 생성')
    parser.add_argument('--output-dataset', help='업데이트된 데이터셋을 저장할 파일 경로 (기본값: 입력 파일 덮어쓰기)')
    parser.add_argument('--output-groups', help='업데이트된 그룹을 저장할 파일 경로')
    
    args = parser.parse_args()
    
    # 그룹 파일 생성 옵션 처리
    if args.create_groups:
        groups_file = args.groups or 'groups.json'
        create_groups_file(groups_file)
        if not args.output_groups:
            args.output_groups = groups_file
    
    # ID 업데이트 및 동기화 실행
    result = update_ids_and_sync(
        args.dataset, 
        args.groups, 
        args.output_dataset, 
        args.output_groups
    )
    
    print("\n=== 업데이트 완료 ===")
    print(f"데이터셋 파일: {args.output_dataset or args.dataset}")
    if args.output_groups:
        print(f"그룹 파일: {args.output_groups}")
    
    # 누락된 ID 경고
    if result['missing_ids']:
        print(f"\n경고: {len(result['missing_ids'])}개의 ID가 그룹에는 있지만 데이터셋에는 없습니다.")
        print(f"누락된 ID: {result['missing_ids']}")
    
    print("\n처리된 ID 수:")
    print(f"데이터셋 ID 수: {len(result['id_mapping'])}")
    print(f"업데이트된 그룹 수: {len(result['updated_groups'])}")

if __name__ == "__main__":
    main()