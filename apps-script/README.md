# Gigo 공용 저장 백엔드

Gigo 대시보드의 마스터 저장값을 여러 브라우저에서 공유하기 위한 Google Apps Script 코드입니다.

1. Google Apps Script 프로젝트를 생성합니다.
2. `apps-script/Code.gs` 내용을 붙여넣습니다.
3. Script Properties에 `GIGO_MASTER_PASSWORD`를 설정합니다. 스프레드시트에 바인딩된 프로젝트라면 `setGigoMasterPassword()`를 실행해도 됩니다.
4. 웹앱으로 배포합니다. 실행 사용자는 본인, 액세스 권한은 대시보드 사용 환경에 맞게 설정합니다.
5. 배포 URL을 루트의 `sync-config.js`의 `endpoint`에 입력합니다.

공용 저장에는 DCM/SAP 원본 파일 전체를 올리지 않고 월별 계산 결과 스냅샷만 저장합니다.
