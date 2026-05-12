const languages = [
  ["ko", "한국어"],
  ["en", "English"],
  ["ja", "日本語"],
  ["zh", "中文"],
  ["es", "Español"],
  ["fr", "Français"],
  ["de", "Deutsch"],
  ["pt-BR", "Português (Brasil)"],
  ["vi", "Tiếng Việt"],
  ["th", "ไทย"],
  ["id", "Bahasa Indonesia"],
  ["ar", "العربية"]
];

const els = {
  targetLanguage: document.querySelector("#targetLanguage"),
  noiseReduction: document.querySelector("#noiseReduction"),
  sourceTranscript: document.querySelector("#sourceTranscript"),
  muteOutput: document.querySelector("#muteOutput"),
  startButton: document.querySelector("#startButton"),
  stopButton: document.querySelector("#stopButton"),
  clearButton: document.querySelector("#clearButton"),
  sessionState: document.querySelector("#sessionState"),
  stateText: document.querySelector("#stateText"),
  sourceText: document.querySelector("#sourceTranscriptText"),
  targetText: document.querySelector("#targetTranscriptText"),
  targetLabel: document.querySelector("#targetLabel"),
  sourceLabel: document.querySelector("#sourceLabel"),
  connectionState: document.querySelector("#connectionState"),
  eventLog: document.querySelector("#eventLog"),
  micMeter: document.querySelector("#micMeter"),
  elapsedTime: document.querySelector("#elapsedTime"),
  audioFrames: document.querySelector("#audioFrames"),
  permissionHelp: document.querySelector("#permissionHelp"),
  permissionHelpText: document.querySelector("#permissionHelpText"),
  retryMicButton: document.querySelector("#retryMicButton")
};

const state = {
  pc: null,
  events: null,
  localStream: null,
  remoteAudio: null,
  audioContext: null,
  meterTimer: null,
  clockTimer: null,
  startedAt: 0,
  audioFrames: 0,
  sourceText: "",
  targetText: "",
  stopping: false
};

init();

function init() {
  for (const [value, label] of languages) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    els.targetLanguage.append(option);
  }

  els.targetLanguage.value = "ko";
  syncLanguageLabel();
  els.targetLanguage.addEventListener("change", () => {
    syncLanguageLabel();
    updateLiveSessionLanguage();
  });
  els.noiseReduction.addEventListener("change", updateLiveNoiseReduction);
  els.sourceTranscript.addEventListener("change", updateLiveSourceTranscript);
  els.muteOutput.addEventListener("change", () => {
    if (state.remoteAudio) {
      state.remoteAudio.muted = els.muteOutput.checked;
    }
  });
  els.startButton.addEventListener("click", startTranslation);
  els.stopButton.addEventListener("click", stopTranslation);
  els.clearButton.addEventListener("click", clearTranscripts);
  els.retryMicButton.addEventListener("click", startTranslation);

  if (!navigator.mediaDevices?.getUserMedia || !window.RTCPeerConnection) {
    setStatus("error", "이 브라우저는 WebRTC 마이크 연결을 지원하지 않습니다.");
    els.startButton.disabled = true;
  }

  checkServerHealth();
}

async function checkServerHealth() {
  try {
    const health = await fetch("/api/health").then((response) => response.json());
    if (!health.hasApiKey) {
      logEvent("OPENAI_API_KEY가 아직 설정되지 않았습니다.");
    }
  } catch {
    logEvent("로컬 서버 상태를 확인하지 못했습니다.");
  }
}

async function startTranslation() {
  if (state.pc) return;
  setControlsBusy(true);
  hidePermissionHelp();
  state.stopping = false;
  state.audioFrames = 0;
  els.audioFrames.textContent = "0";
  setStatus("idle", "세션 준비 중");

  try {
    const { value: clientSecret, session } = await fetchTranslationSecret();
    logEvent(`client secret 발급: ${session?.id || "translation session"}`);

    state.localStream = await navigator.mediaDevices.getUserMedia({
      audio: {
        echoCancellation: true,
        noiseSuppression: true,
        autoGainControl: true
      }
    });

    setupMicMeter(state.localStream);
    state.pc = new RTCPeerConnection();
    state.pc.addTrack(state.localStream.getAudioTracks()[0], state.localStream);

    state.remoteAudio = new Audio();
    state.remoteAudio.autoplay = true;
    state.remoteAudio.playsInline = true;
    state.remoteAudio.muted = els.muteOutput.checked;

    state.pc.ontrack = ({ streams }) => {
      state.remoteAudio.srcObject = streams[0];
      state.remoteAudio.play().catch(() => {
        logEvent("브라우저 자동재생이 막혔습니다. 화면을 한 번 클릭한 뒤 다시 시도하세요.");
      });
    };

    state.pc.onconnectionstatechange = () => {
      const current = state.pc?.connectionState || "closed";
      els.connectionState.textContent = current;
      if (current === "connected") {
        setStatus("live", "통역 중");
      }
      if (["failed", "disconnected", "closed"].includes(current) && !state.stopping) {
        setStatus("error", `연결 상태: ${current}`);
      }
    };

    state.events = state.pc.createDataChannel("oai-events");
    state.events.onopen = () => logEvent("이벤트 채널 연결");
    state.events.onmessage = ({ data }) => handleRealtimeEvent(JSON.parse(data));
    state.events.onerror = () => logEvent("이벤트 채널 오류");
    state.events.onclose = () => logEvent("이벤트 채널 종료");

    const offer = await state.pc.createOffer();
    await state.pc.setLocalDescription(offer);
    await waitForIceGathering(state.pc);

    const sdpResponse = await fetch("https://api.openai.com/v1/realtime/translations/calls", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${clientSecret}`,
        "Content-Type": "application/sdp"
      },
      body: state.pc.localDescription?.sdp || offer.sdp
    });

    if (!sdpResponse.ok) {
      const errorBody = await sdpResponse.text();
      throw new Error(parseOpenAIErrorText(errorBody) || `Realtime call failed (${sdpResponse.status})`);
    }

    await state.pc.setRemoteDescription({
      type: "answer",
      sdp: await sdpResponse.text()
    });

    state.startedAt = Date.now();
    state.clockTimer = window.setInterval(updateElapsedTime, 500);
    setStatus("live", "통역 중");
    setControlsLive(true);
  } catch (error) {
    const message = readableError(error);
    setStatus("error", message);
    logEvent(message);
    if (isMicrophonePermissionError(error)) {
      showPermissionHelp();
    }
    await cleanup();
    setControlsLive(false);
  } finally {
    setControlsBusy(false);
  }
}

async function fetchTranslationSecret() {
  const response = await fetch("/api/translation-session", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      targetLanguage: els.targetLanguage.value,
      noiseReduction: els.noiseReduction.value,
      sourceTranscript: els.sourceTranscript.checked
    })
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(normalizeErrorDetail(data) || "translation session request failed");
  }
  return data;
}

function handleRealtimeEvent(event) {
  if (!event?.type) return;

  if (event.type === "session.input_transcript.delta") {
    if (!els.sourceTranscript.checked) return;
    appendTranscript("source", event.delta);
    return;
  }

  if (event.type === "session.output_transcript.delta") {
    appendTranscript("target", event.delta);
    return;
  }

  if (event.type === "session.output_audio.delta") {
    state.audioFrames += 1;
    els.audioFrames.textContent = String(state.audioFrames);
    return;
  }

  if (event.type === "session.created") {
    logEvent("translation session created");
    return;
  }

  if (event.type === "session.updated") {
    logEvent("session updated");
    return;
  }

  if (event.type === "session.closed") {
    logEvent("session closed");
    void stopTranslation();
    return;
  }

  if (event.type === "error") {
    const message = event.error?.message || "Realtime API error";
    setStatus("error", message);
    logEvent(message);
    return;
  }

  if (!event.type.includes("audio")) {
    logEvent(event.type);
  }
}

function appendTranscript(kind, delta = "") {
  if (!delta) return;
  const key = kind === "source" ? "sourceText" : "targetText";
  const element = kind === "source" ? els.sourceText : els.targetText;
  state[key] += delta;
  element.textContent = state[key];
  element.scrollTop = element.scrollHeight;
}

async function stopTranslation() {
  if (!state.pc && !state.localStream) return;
  state.stopping = true;
  setStatus("idle", "종료 중");
  sendEvent({ type: "session.close", event_id: `close_${Date.now()}` });
  await new Promise((resolve) => window.setTimeout(resolve, 180));
  await cleanup();
  setControlsLive(false);
  setStatus("idle", "대기 중");
  els.connectionState.textContent = "연결 전";
}

async function cleanup() {
  window.clearInterval(state.clockTimer);
  window.clearInterval(state.meterTimer);
  state.clockTimer = null;
  state.meterTimer = null;
  state.startedAt = 0;
  els.micMeter.style.width = "0%";
  els.elapsedTime.textContent = "00:00";

  if (state.events && state.events.readyState !== "closed") {
    state.events.close();
  }

  if (state.pc) {
    state.pc.getSenders().forEach((sender) => sender.track?.stop());
    state.pc.close();
  }

  state.localStream?.getTracks().forEach((track) => track.stop());
  state.remoteAudio?.pause();
  if (state.remoteAudio) {
    state.remoteAudio.srcObject = null;
  }

  if (state.audioContext && state.audioContext.state !== "closed") {
    await state.audioContext.close().catch(() => {});
  }

  state.pc = null;
  state.events = null;
  state.localStream = null;
  state.remoteAudio = null;
  state.audioContext = null;
}

function sendEvent(event) {
  if (state.events?.readyState === "open") {
    state.events.send(JSON.stringify(event));
  }
}

function updateLiveSessionLanguage() {
  if (!state.events || state.events.readyState !== "open") return;
  sendEvent({
    type: "session.update",
    event_id: `language_${Date.now()}`,
    session: {
      audio: {
        output: {
          language: els.targetLanguage.value
        }
      }
    }
  });
}

function updateLiveNoiseReduction() {
  if (!state.events || state.events.readyState !== "open") return;
  const selected = els.noiseReduction.value;
  sendEvent({
    type: "session.update",
    event_id: `noise_${Date.now()}`,
    session: {
      audio: {
        input: {
          noise_reduction: selected === "off" ? null : { type: selected }
        }
      }
    }
  });
}

function updateLiveSourceTranscript() {
  if (!state.events || state.events.readyState !== "open") return;
  if (!els.sourceTranscript.checked) {
    logEvent("원문 자막 표시를 중지했습니다.");
    return;
  }
  sendEvent({
    type: "session.update",
    event_id: `transcript_${Date.now()}`,
    session: {
      audio: {
        input: {
          transcription: { model: "gpt-realtime-whisper" }
        }
      }
    }
  });
}

function clearTranscripts() {
  state.sourceText = "";
  state.targetText = "";
  els.sourceText.innerHTML = '<span class="empty">통역을 시작하면 원문 자막이 여기에 표시됩니다.</span>';
  els.targetText.innerHTML = '<span class="empty">번역 자막이 여기에 누적됩니다.</span>';
  els.eventLog.innerHTML = "";
}

function setupMicMeter(stream) {
  const AudioContext = window.AudioContext || window.webkitAudioContext;
  if (!AudioContext) return;

  state.audioContext = new AudioContext();
  const source = state.audioContext.createMediaStreamSource(stream);
  const analyser = state.audioContext.createAnalyser();
  analyser.fftSize = 512;
  source.connect(analyser);

  const values = new Uint8Array(analyser.frequencyBinCount);
  state.meterTimer = window.setInterval(() => {
    analyser.getByteFrequencyData(values);
    const average = values.reduce((sum, value) => sum + value, 0) / values.length;
    const width = Math.min(100, Math.round(average * 1.8));
    els.micMeter.style.width = `${width}%`;
  }, 100);
}

function updateElapsedTime() {
  if (!state.startedAt) return;
  const seconds = Math.floor((Date.now() - state.startedAt) / 1000);
  const minutes = Math.floor(seconds / 60).toString().padStart(2, "0");
  const rest = (seconds % 60).toString().padStart(2, "0");
  els.elapsedTime.textContent = `${minutes}:${rest}`;
}

function waitForIceGathering(pc) {
  if (pc.iceGatheringState === "complete") {
    return Promise.resolve();
  }
  return new Promise((resolve) => {
    const timeout = window.setTimeout(done, 2500);
    pc.addEventListener("icegatheringstatechange", () => {
      if (pc.iceGatheringState === "complete") done();
    });

    function done() {
      window.clearTimeout(timeout);
      resolve();
    }
  });
}

function setControlsBusy(isBusy) {
  els.startButton.disabled = isBusy || Boolean(state.pc);
  els.stopButton.disabled = isBusy || !state.pc;
}

function setControlsLive(isLive) {
  els.startButton.disabled = isLive;
  els.stopButton.disabled = !isLive;
}

function setStatus(kind, text) {
  els.sessionState.dataset.state = kind;
  els.stateText.textContent = text;
}

function showPermissionHelp() {
  const inAppHint = "Codex 인앱 브라우저에서 계속 거부되면 Edge나 Chrome으로 http://localhost:3000/ 을 열어 주세요.";
  els.permissionHelpText.textContent =
    `주소창 왼쪽의 사이트 정보에서 마이크를 허용한 뒤 새로고침하고 다시 시작하세요. Windows 설정 > 개인정보 및 보안 > 마이크에서도 브라우저 접근이 켜져 있어야 합니다. ${inAppHint}`;
  els.permissionHelp.hidden = false;
}

function hidePermissionHelp() {
  els.permissionHelp.hidden = true;
}

function syncLanguageLabel() {
  const selected = languages.find(([value]) => value === els.targetLanguage.value);
  els.targetLabel.textContent = selected?.[1] || els.targetLanguage.value;
}

function logEvent(message) {
  const item = document.createElement("li");
  const time = new Date().toLocaleTimeString("ko-KR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
  item.textContent = `${time} ${message}`;
  els.eventLog.prepend(item);
  while (els.eventLog.children.length > 40) {
    els.eventLog.lastElementChild?.remove();
  }
}

function readableError(error) {
  const rawMessage = error instanceof Error ? error.message : String(error);
  const message = parseOpenAIErrorText(rawMessage) || rawMessage;
  if (message.includes("OPENAI_API_KEY")) {
    return "OPENAI_API_KEY를 .env에 설정한 뒤 서버를 다시 시작하세요.";
  }
  if (message.includes("insufficient_quota") || message.includes("exceeded your current quota")) {
    return "OpenAI API 사용 한도 또는 결제 크레딧이 부족합니다. OpenAI Platform의 Billing/Usage에서 결제 수단, 크레딧, 프로젝트 한도를 확인해 주세요.";
  }
  if (isMicrophonePermissionError(error)) {
    return "마이크 권한이 거부되었습니다. 사이트 권한 또는 브라우저 마이크 접근을 허용해 주세요.";
  }
  return message;
}

function normalizeErrorDetail(data) {
  if (!data) return "";
  if (typeof data.detail === "string") return data.detail;
  if (typeof data.error === "string") return data.error;
  if (data.error?.message) return data.error.message;
  return parseOpenAIErrorText(JSON.stringify(data));
}

function parseOpenAIErrorText(text) {
  try {
    const parsed = JSON.parse(text);
    if (typeof parsed.detail === "string") return parsed.detail;
    if (typeof parsed.error === "string") return parsed.error;
    if (parsed.error?.message) {
      const code = parsed.error.code ? ` (${parsed.error.code})` : "";
      return `${parsed.error.message}${code}`;
    }
    return "";
  } catch {
    return "";
  }
}

function isMicrophonePermissionError(error) {
  const name = error instanceof DOMException ? error.name : "";
  const message = error instanceof Error ? error.message : String(error);
  return ["NotAllowedError", "PermissionDeniedError", "SecurityError"].includes(name)
    || message.includes("Permission denied")
    || message.includes("NotAllowedError");
}
