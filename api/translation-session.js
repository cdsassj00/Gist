import { createHash } from "node:crypto";
import { requestTranslationClientSecret } from "../lib/realtime-session.mjs";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  if (!process.env.OPENAI_API_KEY) {
    res.status(503).json({
      error: "OPENAI_API_KEY is not set",
      detail: "Add OPENAI_API_KEY to the Vercel project environment variables."
    });
    return;
  }

  try {
    const payload = parseRequestBody(req.body);
    const result = await requestTranslationClientSecret({
      apiKey: process.env.OPENAI_API_KEY,
      payload,
      safetyIdentifier: safetyIdentifier(req)
    });

    res.status(result.status);
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.send(result.text);
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error instanceof Error ? error.message : String(error)
    });
  }
}

function parseRequestBody(body) {
  if (!body) return {};
  if (typeof body === "object") return body;
  if (typeof body === "string") return JSON.parse(body);
  if (Buffer.isBuffer(body)) return JSON.parse(body.toString("utf8"));
  return {};
}

function safetyIdentifier(req) {
  const source = [
    "vercel-realtime-translate",
    req.headers["x-forwarded-for"] || req.socket?.remoteAddress || "unknown",
    req.headers["user-agent"] || "unknown"
  ].join("|");
  return createHash("sha256").update(source).digest("hex");
}
