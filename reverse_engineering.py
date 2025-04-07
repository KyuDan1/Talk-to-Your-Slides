import zipfile
import io
import os
import olefile

def extract_vba_from_ppt(ppt_path, output_dir):
    # PPT 파일이 실제로 .pptm(매크로 지원) 형식인지 확인
    if not ppt_path.endswith(('.pptm', '.ppt')):
        print("이 파일은 매크로를 포함하지 않을 수 있습니다.")
    
    # 출력 디렉토리 생성
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # .pptm 파일을 압축 파일로 처리
    try:
        with zipfile.ZipFile(ppt_path, 'r') as z:
            # vbaProject.bin 파일이 있는지 확인
            if 'ppt/vbaProject.bin' in z.namelist():
                # vbaProject.bin 파일 추출
                z.extract('ppt/vbaProject.bin', output_dir)
                vba_path = os.path.join(output_dir, 'ppt', 'vbaProject.bin')
                
                # olefile을 사용하여 VBA 매크로 추출
                if olefile.isOleFile(vba_path):
                    ole = olefile.OleFile(vba_path)
                    
                    # VBA 모듈 디렉토리 파싱
                    vba_root = 'VBA/'
                    for item in ole.listdir():
                        if item[0].startswith('VBA'):
                            if item[0].endswith('/'):
                                continue
                            
                            vba_filename = item[0].split('/')[-1]
                            if vba_filename.startswith('_'):
                                continue
                                
                            # 모듈 내용 추출
                            vba_content = ole.openstream(item[0]).read().decode('utf-8', errors='ignore')
                            
                            # 파일에 저장
                            with open(os.path.join(output_dir, f"{vba_filename}.bas"), 'w', encoding='utf-8') as f:
                                f.write(vba_content)
                            
                            print(f"추출된 VBA 모듈: {vba_filename}")
                
                print(f"VBA 코드가 {output_dir}에 추출되었습니다.")
            else:
                print("이 PPT 파일에는 VBA 코드가 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")

# 사용 예:
# extract_vba_from_ppt("example.pptm", "output_folder")