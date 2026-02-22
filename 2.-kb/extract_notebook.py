import json
import sys

# Jupyter Notebook에서 Python 코드 추출
def extract_code_from_notebook(notebook_path, output_path):
    try:
        with open(notebook_path, 'r', encoding='utf-8') as f:
            notebook = json.load(f)

        code_cells = []
        for cell in notebook.get('cells', []):
            if cell.get('cell_type') == 'code':
                source = cell.get('source', [])
                if isinstance(source, list):
                    code = ''.join(source)
                else:
                    code = source

                # 빈 셀 제외
                if code.strip():
                    code_cells.append(code)

        # 모든 코드 셀을 합쳐서 저장
        full_code = '\n\n# ' + '='*60 + '\n\n'.join(code_cells)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_code)

        print(f"성공! {len(code_cells)}개의 코드 셀을 추출했습니다.")
        print(f"출력 파일: {output_path}")
        print(f"총 {len(full_code)} 문자")

    except json.JSONDecodeError as e:
        print(f"JSON 파싱 에러: {e}")
        print("파일이 올바른 Jupyter Notebook 형식이 아닙니다.")
    except Exception as e:
        print(f"에러 발생: {e}")

if __name__ == '__main__':
    extract_code_from_notebook(
        'kb-download-20250816.py',
        'kb-download-20250816-extracted.py'
    )
