# GPT Realtime 동시통역 앱

`gpt-realtime-translate`를 쓰는 로컬 브라우저 동시통역 프로그램입니다. 브라우저는 마이크를 WebRTC로 OpenAI Realtime Translation API에 연결하고, Node.js `.mjs` 서버는 표준 API 키 대신 짧게 살아있는 client secret만 발급합니다.

## 왜 `.mjs`인가?

`.mjs`는 Node.js에서 ES Module 파일임을 명시하는 확장자입니다. 이 서버는 `import { createServer } from "node:http"` 같은 ESM 문법을 쓰기 때문에 `.mjs`로 두었습니다. `package.json`에 `"type": "module"`을 넣으면 `.js`로도 가능하지만, 여기서는 서버 파일 성격을 분명히 하려고 `.mjs`를 사용했습니다.

## 실행

1. `.env.example`을 참고해 `C:\_codex\gist\.env` 파일을 만듭니다.

```env
OPENAI_API_KEY=sk-your-openai-api-key
PORT=3000
```

2. Node 서버를 실행합니다.

```powershell
npm start
```

3. 브라우저에서 엽니다.

```text
http://localhost:3000
```

## Vercel 배포

GitHub Pages는 정적 파일만 배포하므로 이 앱의 `/api/translation-session`을 실행할 수 없습니다. Vercel에서는 `api/translation-session.js`가 서버리스 함수로 실행되고, `public/` 폴더의 화면 파일은 정적으로 배포됩니다.

1. GitHub에 올릴 때 `.env`는 올리지 않습니다. `.gitignore`에 이미 제외했습니다.
2. Vercel에서 GitHub 저장소를 Import합니다.
3. Project Settings > Environment Variables에 `OPENAI_API_KEY`를 추가합니다.
4. 빌드 명령은 비워두거나 기본값을 사용해도 됩니다. 이 프로젝트는 별도 빌드가 없습니다.
5. 배포 후 `/api/health`에서 `server: "vercel-node-function"`과 `hasApiKey: true`가 보이면 API 함수가 정상입니다.

## 기능

- 대상 언어 선택
- 브라우저 마이크 입력
- 번역 음성 재생
- 원문 자막과 번역 자막 스트리밍
- 근거리/원거리 마이크 노이즈 보정 설정
- 실행 중 대상 언어와 자막 설정 업데이트

## 파일 구조

- `server.mjs`: Node.js ES Module 기반 로컬 서버와 OpenAI client secret 발급 프록시
- `api/translation-session.js`: Vercel 배포용 client secret 발급 서버리스 함수
- `api/health.js`: Vercel 배포용 상태 확인 함수
- `lib/realtime-session.mjs`: 로컬 서버와 Vercel 함수가 공유하는 OpenAI 세션 생성 로직
- `public/index.html`: 앱 화면
- `public/styles.css`: 반응형 UI 스타일
- `public/app.js`: WebRTC 연결과 실시간 이벤트 처리

## 참고한 공식 문서

- Realtime translation: https://developers.openai.com/api/docs/guides/realtime-translation
- Translation client secret: https://developers.openai.com/api/reference/resources/realtime/subresources/translations/subresources/client_secrets/methods/create
- Translation client/server events: https://developers.openai.com/api/reference/resources/realtime/translation-client-events
- Model card: https://developers.openai.com/api/docs/models/gpt-realtime-translate
