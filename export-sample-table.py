import json

with open('/tmp/sample_results.json', encoding='utf-8') as f:
    data = json.load(f)

headers = ['fileName','name','gender','birth','affiliation','phone','email','education','expertise']
print('| 파일 | 성명 | 성별 | 출생년월일 | 현소속 | 연락처 | 이메일 | 최종학력 | 전문분야 |')
print('| --- | --- | --- | --- | --- | --- | --- | --- | --- |')
for row in data:
    vals = [row.get('fileName',''), row.get('name',''), row.get('gender',''), row.get('birth',''), row.get('affiliation',''), row.get('phone',''), row.get('email',''), row.get('education',''), row.get('expertise','')]
    vals = [str(v).replace('\n','<br>').replace('|','/') for v in vals]
    print('| ' + ' | '.join(vals) + ' |')
