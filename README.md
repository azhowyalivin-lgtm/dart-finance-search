# dart-finance-search

DART(전자공시시스템) Open API로 상장사 재무제표를 검색하고,
서식이 적용된 엑셀로 내려받는 Streamlit 웹앱입니다.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://dartsearch.streamlit.app/)

※ 처음 접속 시 앱을 깨우는 데 30초 정도 걸릴 수 있습니다

## 만든 이유
DART·OpenDART에서 재무제표를 일일이 조회하는 절차가 번거롭고 다운로드 형식도 제각각이라,
재무제표 분석과 회사 간 비교가 쉽도록 엑셀 양식을 하나로 통일하는 도구를 만들었습니다.

## 실행 화면
<img width="1920" height="2297" alt="image" src="https://github.com/user-attachments/assets/6c5828f8-781b-4502-bef5-dec82d84cc14" />


## 주요 기능
- 회사명 검색 → 사업/반기/분기보고서, 연결/별도 재무제표 조회
- 매출액·영업이익·부채비율·영업이익률 요약 지표 (3개월/누적 선택)
- XBRL 재무제표를 서식 포함 엑셀로 다운로드
- 감사보고서(반기·분기는 보고서 본문) 원문에서 재무제표와 주석을 추출해 회사 원본 양식 엑셀로 저장

## 해결한 문제
- **계정 순서 오류:** 일부 공시의 XBRL 데이터가 계정 ID 알파벳순으로 제공되어 순서가 뒤섞이는 문제를 발견하고,
  원문 표의 행 순서에 계정명을 매칭(정확 일치 → 유사도 매칭)해 재정렬
- **반기·분기보고서 처리:** 감사보고서가 없는 반기·분기는 본문 'III. 재무에 관한 사항'에서 표와 주석을 추출하고,
  섹션 경계를 인식해 연결/별도 주석이 섞이지 않도록 처리
- **API 키 보안:** 키를 코드에 넣지 않고 환경변수로 분리

## 사용 기술
Python, Streamlit, OpenDART API, pandas, openpyxl, lxml

## 실행 방법
1. Python 3.10 이상 설치 (https://www.python.org, 설치 시 "Add Python to PATH" 체크)
2. 이 페이지의 초록색 Code 버튼 → Download ZIP → 압축 해제
3. DART API 키 발급 (https://opendart.fss.or.kr, 무료)
4. 압축 푼 폴더에서 명령 프롬프트(cmd)를 열고 아래를 한 줄씩 입력

```
pip install -r requirements.txt
set DART_API_KEY=발급받은키
streamlit run app.py
```

5. 브라우저가 자동으로 열리며 앱이 실행됩니다.

※ Anaconda를 사용한다면 cmd 대신 Anaconda Prompt에서 같은 명령을 입력하세요.
※ 회사 목록(dart.db)은 처음 실행할 때 자동으로 생성됩니다. (1분 정도 소요)
