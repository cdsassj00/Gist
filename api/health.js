import { REALTIME_TRANSLATE_MODEL } from "../lib/realtime-session.mjs";

export default function handler(req, res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  res.status(200).json({
    ok: true,
    model: REALTIME_TRANSLATE_MODEL,
    hasApiKey: Boolean(process.env.OPENAI_API_KEY),
    server: "vercel-node-function"
  });
}
