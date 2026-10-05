# dart-finance-search
searching financial report data 
# dart-finance-search

DART(전자공시시스템) Open API로 상장사 재무제표를 검색하고,
서식이 적용된 엑셀로 내려받는 Streamlit 웹앱입니다.

## 주요 기능
- 회사명 검색 → 사업/반기/분기보고서, 연결/별도 재무제표 조회
- 매출액·영업이익·부채비율·영업이익률 요약 지표 (3개월/누적 선택)
- XBRL 재무제표를 서식 포함 엑셀로 다운로드
- 감사보고서(반기·분기는 보고서 본문) 원문을 파싱해 재무제표 + 주석을 회사 원본 양식 엑셀로 저장

## 해결한 문제
- 일부 공시의 XBRL 데이터가 계정 ID 알파벳순으로 제공되어 순서가 뒤섞이는 문제를 발견하고
  원문 표의 행 순서에 계정명을 매칭(정확 일치 → 유사도 매칭)해 재정렬
- 반기·분기보고서는 감사보고서가 없어 본문 'III. 재무에 관한 사항'을 파싱하고(원문을 읽어 표와 주석을 추출해)
  섹션 경계를 인식해 연결/별도 주석이 섞이지 않도록 처리

## 실행 방법
pip install -r requirements.txt
DART_API_KEY 환경변수 설정 후
streamlit run app.py

※ 회사 목록(dart.db)은 DART 고유번호 파일(corpCode.xml)로 생성합니다.
