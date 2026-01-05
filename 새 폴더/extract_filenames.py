import re

# JSON 파일 읽기
with open('list.json', 'r', encoding='utf-8') as f:
    content = f.read()

# filename 추출 (정규표현식 사용)
pattern = r'"filename":\s*"([^"]+)"'
filenames = re.findall(pattern, content)

# 결과 출력
print(f"총 {len(filenames)}개의 filename을 추출했습니다:\n")
for filename in filenames:
    print(filename)

# 파일로 저장
with open('filenames.txt', 'w', encoding='utf-8') as f:
    for filename in filenames:
        f.write(filename + '\n')

print(f"\n결과가 'filenames.txt' 파일에 저장되었습니다.")

