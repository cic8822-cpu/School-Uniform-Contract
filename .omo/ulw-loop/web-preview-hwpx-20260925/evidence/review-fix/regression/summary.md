# 회귀 검증

- `webapp`: `npm run lint` → 종료 코드 0. 실행 로그: `npm-lint-final.log`.
- `webapp`: `npm run build` → 종료 코드 0. 실행 로그: `npm-build-final.log`.
- `webapp`: `node scripts/verify-hwpx-preview-cache.mjs` → `{"status":"PASS","fetchCalls":2}`. 실행 로그: `../green/hwpx-cache-retry-final.log`.
- 저장소 루트: `git diff --check` → 종료 코드 0. 실행 로그: `git-diff-check-final.log`.
