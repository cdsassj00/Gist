import { requestTranslationClientSecret } from "../../lib/realtime-session.mjs";

declare const Netlify: {
  env: {
    get(name: string): string | undefined;
  };
};

export default async (req: Request) => {
  if (req.method !== "POST") {
    return Response.json({ error: "Method not allowed" }, { status: 405 });
  }

  const apiKey = Netlify.env.get("OPENAI_API_KEY");
  if (!apiKey) {
    return Response.json(
      {
        error: "OPENAI_API_KEY is not set",
        detail: "Add OPENAI_API_KEY to the Netlify site environment variables."
      },
      { status: 503 }
    );
  }

  try {
    const payload = await req.json().catch(() => ({}));
    const result = await requestTranslationClientSecret({
      apiKey,
      payload,
      safetyIdentifier: await safetyIdentifier(req)
    });

    return new Response(result.text, {
      status: result.status,
      headers: {
        "Content-Type": "application/json; charset=utf-8"
      }
    });
  } catch (error) {
    return Response.json(
      {
        error: "Server error",
        detail: error instanceof Error ? error.message : String(error)
      },
      { status: 500 }
    );
  }
};

export const config = {
  path: "/api/translation-session"
};

async function safetyIdentifier(req: Request) {
  const source = [
    "netlify-realtime-translate",
    req.headers.get("x-forwarded-for") || "unknown",
    req.headers.get("user-agent") || "unknown"
  ].join("|");
  const bytes = new TextEncoder().encode(source);
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(hash))
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("");
}
