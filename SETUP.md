# GitHub Actions로 PC 없이 매주 실행하기

매주 월요일 09:00(KST)에 GitHub 서버가 수집 스크립트를 실행하고 엑셀을 첨부해 메일로 보냅니다.
`output/seen.json`, `output/last_run.json`은 실행이 끝날 때마다 저장소에 자동 커밋되므로, 이미 보고한 항목은 다음 주에 다시 나오지 않습니다.

## 1. 저장소 만들기 (Private)
1. https://github.com/new 에서 **Private** 저장소를 만듭니다(예: `mcee-waste-weekly`). README 등 초기 파일은 추가하지 않습니다.
2. 이 폴더에서 아래를 실행합니다.
   ```
   git remote add origin https://github.com/<계정>/mcee-waste-weekly.git
   git push -u origin main
   ```

## 2. Secrets 등록
저장소 → Settings → Secrets and variables → Actions → **New repository secret**

| 이름 | 값 |
|---|---|
| `SMTP_USER` | 보내는 Gmail 주소 |
| `SMTP_PASSWORD` | Gmail 앱 비밀번호(16자리) |
| `MAIL_TO` | 받는 주소 |

## 3. 테스트 실행
Actions 탭 → `mcee-waste-weekly` → **Run workflow**.
로그에 "메일 발송 완료"가 나오면 끝입니다. 엑셀과 첨부파일은 실행 화면 하단 Artifacts에서도 90일간 받을 수 있습니다.

## 참고
- 저장소에 60일 이상 커밋이 없으면 GitHub가 예약 실행을 멈추지만, 매주 상태 파일을 커밋하므로 보통 해당되지 않습니다.
- 예약 실행은 GitHub 부하에 따라 몇 분~수십 분 늦게 시작될 수 있습니다.
- 클라우드 실행이 확인되면 PC의 Claude 예약 작업(`mcee-waste-weekly`)을 꺼서 메일 중복을 막습니다.
