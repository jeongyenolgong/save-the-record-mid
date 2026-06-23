# Save the Record — 중등용

기록문화 세계기록유산 카드게임 (중등용, 한국 기록유산 14건).

## 구조 (이 폴더 = 깃헙 레포 1개, GitHub Pages 루트)
```
index.html      게임 본체 (content.json 을 fetch)
content.json    게임 데이터 (convert.py 산출물)
images/         유산 사진 (no 와 동일 파일명 — 확정 후 01.jpg ~ 14.jpg 로 통일 예정)
content/        콘텐츠 원본 엑셀 (작성·검토용)
convert.py      엑셀 → content.json 변환기
```

## 콘텐츠 수정 → 반영
1. `content/기록문화_중등_콘텐츠.xlsx` 에서 수정 (items / content 시트)
2. `python3 convert.py` 실행 → `content.json` 재생성
3. 커밋 & 푸시 → GitHub Pages 자동 반영

## 초등용과의 관계
- **완전 독립 레포.** 번호·정답코드·이미지 모두 초등과 무관하게 자체 관리.
- 본문은 14건 작성 완료(임포트됨). 남은 작업: 신규 13·14(승정원 일기·조선왕조 의궤) 난이도·desc, 제공방식/정답코드 확정, 이미지 파일명 통일.

## 로컬 확인
```
python3 -m http.server 8000
# http://localhost:8000/  접속
```
※ `file://` 금지, HTTP 서버 필수.
