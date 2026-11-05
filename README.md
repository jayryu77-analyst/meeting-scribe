# NDR 미팅노트

휴대폰·PC 브라우저에서 쓰는 NDR(기관 IR 미팅) 일정·노트 웹앱입니다.

- 엑셀(.xlsx)·CSV·텍스트로 일정을 올리면 날짜별 일정표가 만들어집니다.
- 미팅마다 실시간 받아쓰기(음성 파일은 저장하지 않음), 메모, AI 노트 정리
- 하루 미팅이 끝나면 AI 하루 요약 (설정에 Claude API 키 입력 시 자동)
- 모든 데이터는 이 기기 브라우저에만 저장 → 백업 파일로 옮기기
- 주소: https://jayryu77-analyst.github.io/meeting-scribe/

## 파일
| 파일 | 역할 |
|---|---|
| index.html | 화면 구조 |
| style.css | 디자인 |
| app.js | 일정 화면·받아쓰기·AI 정리·백업 |
| parser.js | 엑셀/텍스트 일정 읽기 |
| sw.js, manifest.webmanifest, icon-*.png | 홈 화면 앱 설정 |
