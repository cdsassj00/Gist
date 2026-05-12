export const REALTIME_TRANSLATE_MODEL = "gpt-realtime-translate";

export async function requestTranslationClientSecret({ apiKey, payload, safetyIdentifier }) {
  const targetLanguage = normalizeLanguage(payload.targetLanguage || "ko");
  const transcriptEnabled = payload.sourceTranscript !== false;
  const noiseReduction = normalizeNoiseReduction(payload.noiseReduction || "near_field");

  const inputAudio = {
    transcription: transcriptEnabled ? { model: "gpt-realtime-whisper" } : undefined,
    noise_reduction: noiseReduction
  };

  const body = {
    expires_after: {
      anchor: "created_at",
      seconds: 600
    },
    session: {
      model: REALTIME_TRANSLATE_MODEL,
      audio: {
        input: compact(inputAudio),
        output: {
          language: targetLanguage
        }
      }
    }
  };

  const response = await fetch("https://api.openai.com/v1/realtime/translations/client_secrets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "OpenAI-Safety-Identifier": safetyIdentifier
    },
    body: JSON.stringify(body)
  });

  const text = await response.text();
  if (!response.ok) {
    return {
      status: response.status,
      text: JSON.stringify({
        error: "OpenAI translation session request failed",
        status: response.status,
        detail: parseOpenAIError(text)
      })
    };
  }

  return {
    status: response.status,
    text
  };
}

export function normalizeLanguage(value) {
  const language = String(value).trim();
  const allowed = new Set(["ko", "en", "ja", "zh", "es", "fr", "de", "pt-BR", "vi", "th", "id", "ar"]);
  return allowed.has(language) ? language : "ko";
}

export function normalizeNoiseReduction(value) {
  if (value === "off" || value === null) {
    return null;
  }
  return value === "far_field" ? { type: "far_field" } : { type: "near_field" };
}

export function parseOpenAIError(text) {
  try {
    const data = JSON.parse(text);
    if (data.error?.message) {
      const code = data.error.code ? ` (${data.error.code})` : "";
      return `${data.error.message}${code}`;
    }
    return data.error || data;
  } catch {
    return text;
  }
}

function compact(value) {
  return Object.fromEntries(
    Object.entries(value).filter(([, entry]) => entry !== undefined)
  );
}
