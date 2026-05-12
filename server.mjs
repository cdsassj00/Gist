import { createHash } from "node:crypto";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { readFile } from "node:fs/promises";
import {
  REALTIME_TRANSLATE_MODEL,
  requestTranslationClientSecret
} from "./lib/realtime-session.mjs";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const publicDir = resolve(__dirname, "public");

await loadDotEnv();

const configuredPort = Number(process.env.PORT || 3000);
const port = Number.isFinite(configuredPort) ? configuredPort : 3000;

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".svg", "image/svg+xml"],
  [".ico", "image/x-icon"]
]);

const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", `http://${req.headers.host ?? "localhost"}`);

    if (req.method === "GET" && url.pathname === "/api/health") {
      await loadDotEnv();
      sendJson(res, 200, {
        ok: true,
        model: REALTIME_TRANSLATE_MODEL,
        hasApiKey: Boolean(process.env.OPENAI_API_KEY),
        server: "node-mjs"
      });
      return;
    }

    if (req.method === "POST" && url.pathname === "/api/translation-session") {
      await createTranslationSession(req, res);
      return;
    }

    if (req.method === "GET" || req.method === "HEAD") {
      await serveStatic(url.pathname, req, res);
      return;
    }

    sendJson(res, 405, { error: "Method not allowed" });
  } catch (error) {
    console.error(error);
    sendJson(res, 500, {
      error: "Server error",
      detail: error instanceof Error ? error.message : String(error)
    });
  }
});

server.listen(port, () => {
  console.log(`Realtime interpreter running at http://localhost:${port}`);
  console.log("Server runtime: Node.js ES module (.mjs)");
  if (!process.env.OPENAI_API_KEY) {
    console.log("OPENAI_API_KEY is not set. Add it to .env before starting a live session.");
  }
});

async function createTranslationSession(req, res) {
  await loadDotEnv();

  if (!process.env.OPENAI_API_KEY) {
    sendJson(res, 503, {
      error: "OPENAI_API_KEY is not set",
      detail: "Add OPENAI_API_KEY to .env or your shell environment, then retry."
    });
    return;
  }

  const payload = await readJsonBody(req, 64 * 1024);
  const result = await requestTranslationClientSecret({
    apiKey: process.env.OPENAI_API_KEY,
    payload,
    safetyIdentifier: safetyIdentifier(req)
  });

  res.statusCode = result.status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(result.text);
}

async function serveStatic(pathname, req, res) {
  const requested = pathname === "/" ? "/index.html" : decodeURIComponent(pathname);
  const normalized = normalize(requested).replace(/^(\.\.[/\\])+/, "");
  const filePath = resolve(join(publicDir, normalized));

  if (!filePath.startsWith(publicDir)) {
    sendJson(res, 403, { error: "Forbidden" });
    return;
  }

  try {
    const file = await readFile(filePath);
    const contentType = mimeTypes.get(extname(filePath)) || "application/octet-stream";
    res.statusCode = 200;
    res.setHeader("Content-Type", contentType);
    res.setHeader("Cache-Control", "no-store");
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    res.end(file);
  } catch {
    sendJson(res, 404, { error: "Not found" });
  }
}

async function readJsonBody(req, limit) {
  let body = "";
  for await (const chunk of req) {
    body += chunk;
    if (body.length > limit) {
      throw new Error("Request body is too large");
    }
  }

  if (!body.trim()) {
    return {};
  }

  try {
    return JSON.parse(body);
  } catch {
    throw new Error("Invalid JSON body");
  }
}

function sendJson(res, status, data) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(data));
}

function safetyIdentifier(req) {
  const source = [
    "local-realtime-translate",
    req.socket.remoteAddress || "unknown",
    req.headers["user-agent"] || "unknown"
  ].join("|");
  return createHash("sha256").update(source).digest("hex");
}

async function loadDotEnv() {
  try {
    const envPath = resolve(__dirname, ".env");
    const envFile = await readFile(envPath, "utf8");
    for (const line of envFile.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;
      const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)=(.*)$/);
      if (!match) continue;
      const [, key, rawValue] = match;
      if (process.env[key]) continue;
      process.env[key] = rawValue.replace(/^['"]|['"]$/g, "");
    }
  } catch {
    // .env is optional; shell environment variables work too.
  }
}
