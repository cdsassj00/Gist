import { REALTIME_TRANSLATE_MODEL } from "../../lib/realtime-session.mjs";

declare const Netlify: {
  env: {
    get(name: string): string | undefined;
  };
};

export default async (req: Request) => {
  if (req.method !== "GET") {
    return Response.json({ error: "Method not allowed" }, { status: 405 });
  }

  return Response.json({
    ok: true,
    model: REALTIME_TRANSLATE_MODEL,
    hasApiKey: Boolean(Netlify.env.get("OPENAI_API_KEY")),
    server: "netlify-function"
  });
};

export const config = {
  path: "/api/health"
};
