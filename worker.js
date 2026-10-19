import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/* Settings                                                            */
/* ------------------------------------------------------------------ */

const TEXT_MODEL = "@cf/google/gemma-4-26b-a4b-it";
// Used only when the user sends a photo. Check the Cloudflare model page:
// Llama 3.2 Vision needs a one-time license acceptance.
const VISION_MODEL = "@cf/meta/llama-3.2-11b-vision-instruct";

// Daily request limits per plan. Change the numbers here only.
const DAILY_LIMITS = {
  free: 8,
  premium: 100,
  pro: 300
};

const PRICES = {
  premium: { monthly: 2500, annual: 25000 },
  pro: { monthly: 10000, annual: 110000 }
};

const taskInstructions = {
  correct:
    "Correct grammar, spelling, punctuation, and sentence errors. Preserve the writer's meaning, facts, and voice. Return the corrected text first, then briefly explain the important corrections made.",
  improve:
    "Improve clarity, grammar, flow, and readability while preserving the writer's meaning, facts, and voice. Return the improved text first, then explain the most important improvements.",
  generate:
    "Write original, clear, useful content based on the user's instructions. Follow the requested format and tone. Explain difficult ideas simply when appropriate.",
  explain:
    "Teach the material step by step using simple language. Explain what it means, why it works, and how to apply it. Give a relevant example where helpful. Do not merely repeat the material.",
  summarize:
    "Summarize the supplied material accurately. Preserve its main ideas and important details. Do not invent information. Explain unfamiliar terms briefly when useful.",
  tone:
    "Rewrite the supplied text in the requested tone while preserving its meaning and important details. Briefly explain significant tone changes.",
  rewrite:
    "Rewrite the supplied text to improve clarity and natural flow while preserving its meaning and facts. Briefly explain important changes.",
  expand:
    "Expand the supplied text with relevant explanations, examples, and details. Do not invent facts or citations. Make the added explanations clear and useful."
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization"
};

/* ------------------------------------------------------------------ */
/* Basic helpers                                                       */
/* ------------------------------------------------------------------ */

function respond(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json; charset=utf-8"
    }
  });
}

function extractText(data) {
  if (typeof data === "string") return data.trim();

  const text =
    data?.response ??
    data?.result ??
    data?.choices?.[0]?.message?.content ??
    data?.output?.[0]?.content?.[0]?.text;

  return typeof text === "string" ? text.trim() : "";
}

function stripCodeFences(text) {
  return text
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "");
}

function parseJsonObject(text, failMessage) {
  const cleaned = stripCodeFences(text);

  try {
    return JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");

    if (start < 0 || end <= start) {
      throw new Error(failMessage);
    }

    try {
      return JSON.parse(cleaned.slice(start, end + 1));
    } catch {
      throw new Error(failMessage);
    }
  }
}

function getErrorStatus(error) {
  if ([400, 401, 403, 429, 503].includes(error.status)) {
    return error.status;
  }
  return 500;
}

/* ------------------------------------------------------------------ */
/* Quiz parsing                                                        */
/* ------------------------------------------------------------------ */

function parseQuiz(text) {
  const parsed = parseJsonObject(
    text,
    "The AI did not return valid quiz JSON. Please try again."
  );

  const questions = Array.isArray(parsed) ? parsed : parsed.questions;

  if (!Array.isArray(questions) || !questions.length) {
    throw new Error("No quiz questions were returned. Please try again.");
  }

  const validated = questions.map((item, index) => {
    const question = String(item.question || item.prompt || "").trim();
    const options = item.options || item.choices;

    if (!question || !Array.isArray(options) || options.length < 2) {
      throw new Error(
        `Quiz question ${index + 1} is incomplete. Please generate the quiz again.`
      );
    }

    const cleanOptions = options.map(option =>
      String(
        typeof option === "object"
          ? option.text ?? option.label ?? ""
          : option
      ).trim()
    );

    let answer =
      item.answerIndex ??
      item.correctIndex ??
      item.correctAnswerIndex ??
      item.correctAnswer ??
      item.answer;

    if (typeof answer === "string") {
      const value = answer.trim();
      const letter = value.toUpperCase().match(/^[A-D]$/);

      if (letter) {
        answer = letter[0].charCodeAt(0) - 65;
      } else {
        answer = cleanOptions.findIndex(
          option => option.toLowerCase() === value.toLowerCase()
        );
      }
    }

    answer = Number(answer);

    if (
      !Number.isInteger(answer) ||
      answer < 0 ||
      answer >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid answer. Please try again.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex: answer,
      explanation: String(item.explanation || item.reason || "").trim()
    };
  });

  return { questions: validated };
}

/* ------------------------------------------------------------------ */
/* Excel building                                                      */
/* ------------------------------------------------------------------ */

// Blocks formulas that can reach outside the workbook.
const UNSAFE_FORMULA = /WEBSERVICE|HYPERLINK|\bCALL\b|EXEC|DDE|\||\bREGISTER\b/i;

function buildWorkbook(spec) {
  const sheets = Array.isArray(spec?.sheets) ? spec.sheets.slice(0, 10) : [];

  if (!sheets.length) {
    throw new Error("No spreadsheet data was returned. Please try again.");
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  sheets.forEach((sheet, i) => {
    let name =
      String(sheet?.name || `Sheet${i + 1}`)
        .replace(/[\\/?*[\]:]/g, "")
        .slice(0, 31) || `Sheet${i + 1}`;

    if (usedNames.has(name.toLowerCase())) {
      name = `Sheet${i + 1}`;
    }
    usedNames.add(name.toLowerCase());

    const rows = (Array.isArray(sheet?.rows) ? sheet.rows : [])
      .slice(0, 2000)
      .map(row =>
        (Array.isArray(row) ? row : [row]).slice(0, 50).map(cell => {
          if (cell === null || cell === undefined) return "";
          if (typeof cell === "number") {
            return Number.isFinite(cell) ? cell : "";
          }
          if (typeof cell === "boolean") return cell;

          const value = String(cell);

          if (value.startsWith("=")) {
            if (UNSAFE_FORMULA.test(value) || value.length > 500) {
              // Store as plain text instead of a live formula.
              return { t: "s", v: value };
            }
            return { t: "n", f: value.slice(1) };
          }

          return value;
        })
      );

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Auto column widths
    const widths = [];
    rows.forEach(row =>
      row.forEach((cell, j) => {
        const len =
          cell && typeof cell === "object" ? 12 : String(cell).length;
        widths[j] = Math.max(widths[j] || 8, Math.min(len + 2, 50));
      })
    );
    ws["!cols"] = widths.map(wch => ({ wch }));

    XLSX.utils.book_append_sheet(wb, ws, name);
  });

  return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
}

/* ------------------------------------------------------------------ */
/* AI call                                                             */
/* ------------------------------------------------------------------ */

async function callAI(env, prompt, jsonMode = false, imageDataUrl = null) {
  if (!env.AI || typeof env.AI.run !== "function") {
    const error = new Error(
      "Cloudflare Workers AI is not connected. Check that your Worker has an AI binding named AI."
    );
    error.status = 500;
    throw error;
  }

  const hasImage = Boolean(imageDataUrl);
  const model = hasImage ? VISION_MODEL : TEXT_MODEL;

  const identityRules = [
    "You are Malie AI, the AI assistant created and developed by Campus Creative, founded by Osume Prosper.",
    "IDENTITY RULE: When asked who created you, who made you, who developed you, who owns Malie AI, or who your founder is, answer clearly that Malie AI was created and developed by Campus Creative, founded by Osume Prosper.",
    "Never answer that Google, Gemini, Google AI, Cloudflare, or another technology provider created Malie AI. They may provide underlying technology, but Campus Creative is the company behind the Malie AI product.",
    "If asked who provides your underlying AI technology, explain that Malie AI currently runs on Cloudflare Workers AI and uses the configured Gemma model.",
    "Be truthful, consistent, and confident about your identity. Do not confuse the creator of Malie AI with the provider of its underlying AI model."
  ];

  const tutorRules = [
    "You are Malie AI, a helpful, patient, accurate AI tutor and writing assistant.",
    "Teach rather than merely give answers.",
    "When answering educational questions, explain the reasoning step by step in simple language.",
    "Define unfamiliar terms and use a relevant example when it helps.",
    "For calculations, show the method, the working, and the final answer with units where appropriate.",
    "For science questions, explain why the answer is correct, not just what the answer is.",
    "For writing tasks, follow the requested task and preserve the user's meaning and facts.",
    "Use headings or numbered steps when they improve clarity.",
    "If the request is unclear, ask a brief clarifying question rather than inventing details.",
    "Be honest about uncertainty and do not invent facts, sources, or citations."
  ];

  const systemPrompt = (jsonMode
    ? identityRules
    : [...identityRules, ...tutorRules]
  ).join(" ");

  const userContent = hasImage
    ? [
        { type: "text", text: prompt },
        { type: "image_url", image_url: { url: imageDataUrl } }
      ]
    : prompt;

  const payload = {
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent }
    ],
    max_tokens: jsonMode ? 4096 : 2048
  };

  // These options are only sent to the Gemma text model.
  if (!hasImage) {
    if (jsonMode) payload.response_format = { type: "json_object" };
    payload.chat_template_kwargs = { enable_thinking: false };
  }

  let result;

  try {
    result = await env.AI.run(model, payload);
  } catch (cause) {
    console.error("Workers AI error:", String(cause));

    const message = String(cause?.message || cause);
    const isLimit = /limit|quota|neuron|429/i.test(message);
    const isBusy = /capacity|busy|overload|503/i.test(message);
    const needsLicense = /agree|license/i.test(message) && hasImage;

    const error = new Error(
      needsLicense
        ? "The image model has not been enabled yet. The site owner must accept its license once."
        : isLimit
        ? "The Cloudflare AI usage limit may have been reached. Please try again after the usage limit resets."
        : isBusy
        ? "Cloudflare AI is busy right now. Please wait a little and try again."
        : "Cloudflare AI could not complete the request. Please try again."
    );

    error.status = isLimit ? 429 : 503;
    throw error;
  }

  const text = extractText(result);

  if (!text) {
    console.error(
      "Unexpected Workers AI response:",
      JSON.stringify(result).slice(0, 1000)
    );
    throw new Error("The AI returned no text. Please try again.");
  }

  return text;
}

/* ------------------------------------------------------------------ */
/* Auth, plans, usage                                                  */
/* ------------------------------------------------------------------ */

async function authenticateFirebase(request, env) {
  if (!env.FIREBASE_WEB_API_KEY) {
    const error = new Error("Firebase verification is not configured.");
    error.status = 500;
    throw error;
  }

  const authorization = request.headers.get("Authorization") || "";
  const match = authorization.match(/^Bearer\s+(.+)$/i);

  if (!match) {
    const error = new Error("Please sign in with Google to use Malie AI.");
    error.status = 401;
    throw error;
  }

  let response;

  try {
    response = await fetch(
      "https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" +
        encodeURIComponent(env.FIREBASE_WEB_API_KEY),
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idToken: match[1] })
      }
    );
  } catch {
    const error = new Error("Could not verify your sign-in. Please try again.");
    error.status = 503;
    throw error;
  }

  if (!response.ok) {
    const error = new Error(
      "Your sign-in session is invalid or expired. Please sign in again."
    );
    error.status = 401;
    throw error;
  }

  const data = await response.json();
  const user = data.users?.[0];

  if (!user?.localId) {
    const error = new Error("Could not verify your sign-in.");
    error.status = 401;
    throw error;
  }

  return user.localId;
}

function getUsageDate() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Africa/Lagos",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());

  const values = {};

  for (const part of parts) {
    if (part.type !== "literal") {
      values[part.type] = part.value;
    }
  }

  return `${values.year}-${values.month}-${values.day}`;
}

// Returns "free", "premium" or "pro".
// A paid plan counts only if it is active and has not expired.
async function getActivePlan(env, userId) {
  if (!env.DB) {
    const error = new Error(
      "Usage tracking is not connected. Please try again later."
    );
    error.status = 500;
    throw error;
  }

  const subscription = await env.DB.prepare(
    `SELECT plan, status, current_period_end
     FROM subscriptions
     WHERE user_id = ?
     LIMIT 1`
  ).bind(userId).first();

  if (!subscription) return "free";

  const plan = String(subscription.plan || "").toLowerCase();
  const status = String(subscription.status || "").toLowerCase();
  const expiry = subscription.current_period_end
    ? Date.parse(subscription.current_period_end)
    : null;

  const active =
    status === "active" &&
    expiry !== null &&
    Number.isFinite(expiry) &&
    expiry > Date.now();

  if (active && (plan === "premium" || plan === "pro")) {
    return plan;
  }

  return "free";
}

async function reserveDailyUsage(env, userId, knownPlan = null) {
  const plan = knownPlan || (await getActivePlan(env, userId));
  const limit = DAILY_LIMITS[plan] ?? DAILY_LIMITS.free;
  const usageDate = getUsageDate();

  const result = await env.DB.prepare(
    `INSERT INTO daily_usage (user_id, usage_date, used)
     VALUES (?, ?, 1)
     ON CONFLICT(user_id, usage_date)
     DO UPDATE SET used = used + 1
     WHERE used < ?
     RETURNING used`
  ).bind(userId, usageDate, limit).first();

  if (!result) {
    const error = new Error(
      `You've reached your ${limit} daily requests on the ${plan} plan. ${
        plan === "free"
          ? "Upgrade to Premium or Pro for more daily requests."
          : plan === "premium"
          ? "Upgrade to Pro for more daily requests, or try again after your limit resets tomorrow."
          : "Your daily limit will reset tomorrow."
      }`
    );
    error.status = 429;
    throw error;
  }

  return {
    plan,
    used: result.used,
    limit,
    remaining: limit - result.used,
    date: usageDate
  };
}

// Gives back one request when the AI fails, so users aren't charged for errors.
async function refundDailyUsage(env, userId, usageDate) {
  try {
    await env.DB.prepare(
      `UPDATE daily_usage
       SET used = CASE WHEN used > 0 THEN used - 1 ELSE 0 END
       WHERE user_id = ? AND usage_date = ?`
    ).bind(userId, usageDate).run();
  } catch (error) {
    console.error("Usage refund failed:", String(error));
  }
}

/* ------------------------------------------------------------------ */
/* Malie Video Studio (Pro) — settings and helpers                     */
/* ------------------------------------------------------------------ */

// >>> DAILY VIDEO ALLOWANCE <<<
// Two limits apply to each Pro user per day. A video is refused if EITHER is used up:
//   1. at most VIDEO_DEFAULT_DAILY_VIDEOS videos, and
//   2. at most VIDEO_DEFAULT_DAILY_SECONDS seconds of video in total.
// Each video is 5 or 10 seconds. Worst-case spend per user per day =
// seconds limit x $0.05 (480p) = 15 x 0.05 = $0.75.
// Change the numbers below, or set Worker variables VIDEO_DAILY_VIDEOS and
// VIDEO_DAILY_SECONDS (Settings > Variables) to override them without editing code.
const VIDEO_DEFAULT_DAILY_VIDEOS = 3;
const VIDEO_DEFAULT_DAILY_SECONDS = 15;

const VIDEO_STUDIO = {
  model: "xai/grok-imagine-video",
  resolution: "480p",
  durations: [5, 10],
  aspectRatios: ["16:9", "9:16", "1:1"],
  promptMin: 10,
  promptMax: 800,
  timeoutMs: 170000
};

function videoError(message, status, refund = true) {
  const error = new Error(message);
  error.status = status;
  error.refund = refund;
  return error;
}

function getVideoDailyLimit(env) {
  const raw = Number.parseInt(env.VIDEO_DAILY_VIDEOS, 10);
  return Number.isInteger(raw) && raw >= 0 && raw <= 50
    ? raw
    : VIDEO_DEFAULT_DAILY_VIDEOS;
}

function getVideoDailySeconds(env) {
  const raw = Number.parseInt(env.VIDEO_DAILY_SECONDS, 10);
  return Number.isInteger(raw) && raw >= 0 && raw <= 600
    ? raw
    : VIDEO_DEFAULT_DAILY_SECONDS;
}

function videoUsageSummary(limit, videos, seconds, secondsLimit, date) {
  return {
    limitVideos: limit,
    usedVideos: videos,
    remainingVideos: Math.max(0, limit - videos),
    limitSeconds: secondsLimit,
    usedSeconds: seconds,
    remainingSeconds: Math.max(0, secondsLimit - seconds),
    date
  };
}

async function getVideoUsage(env, userId, usageDate) {
  const row = await env.DB.prepare(
    `SELECT videos_used, seconds_used FROM video_usage
     WHERE user_id = ? AND usage_date = ?`
  ).bind(userId, usageDate).first();

  return {
    videos: Number(row?.videos_used || 0),
    seconds: Number(row?.seconds_used || 0)
  };
}

// Atomically counts one video (and its seconds) for today, but only if the
// user is still under the daily limit. Two simultaneous requests cannot both
// take the last slot.
async function reserveVideo(env, userId, usageDate, seconds, limit, secondsLimit) {
  const limitError = () =>
    videoError(
      limit <= 0 || secondsLimit <= 0
        ? "Video generation is switched off right now."
        : `You don't have enough of today's video allowance left for this video (daily limit: ${limit} video${limit === 1 ? "" : "s"} and ${secondsLimit} seconds). Try a shorter video, or come back after midnight (Lagos time).`,
      429,
      false
    );

  if (limit <= 0 || secondsLimit <= 0 || seconds > secondsLimit) throw limitError();

  const row = await env.DB.prepare(
    `INSERT INTO video_usage (user_id, usage_date, seconds_used, videos_used)
     VALUES (?, ?, ?, 1)
     ON CONFLICT(user_id, usage_date)
     DO UPDATE SET
       seconds_used = seconds_used + excluded.seconds_used,
       videos_used = videos_used + 1
     WHERE videos_used < ?
       AND seconds_used + excluded.seconds_used <= ?
     RETURNING videos_used, seconds_used`
  ).bind(userId, usageDate, seconds, limit, secondsLimit).first();

  if (!row) throw limitError();

  return { videos: Number(row.videos_used), seconds: Number(row.seconds_used) };
}

async function refundVideo(env, userId, usageDate, seconds) {
  try {
    await env.DB.prepare(
      `UPDATE video_usage
       SET seconds_used = MAX(seconds_used - ?, 0),
           videos_used = MAX(videos_used - 1, 0)
       WHERE user_id = ? AND usage_date = ?`
    ).bind(seconds, userId, usageDate).run();
  } catch (error) {
    console.error("Video usage refund failed:", String(error));
  }
}

// Cloudflare documents the result as { state: "Completed", result: { video: "<url>" } }
// and the raw output schema as { video: "<url>" }. Both shapes are accepted.
function extractVideoUrl(result) {
  const state = typeof result?.state === "string" ? result.state.toLowerCase() : "";

  if (state && !["completed", "complete", "succeeded", "success"].includes(state)) {
    if (/fail|error|cancel|reject/.test(state)) {
      throw videoError("The video could not be generated. Please try again.", 502, true);
    }
    // Not finished and no way to check again: do not claim success, do not refund.
    throw videoError(
      "The video provider has not confirmed completion, so no video is available. Please wait a minute before trying again.",
      504,
      false
    );
  }

  const candidate =
    typeof result === "string"
      ? result
      : result?.result?.video ?? result?.video ?? result?.result?.video?.url ?? result?.video?.url;

  if (typeof candidate !== "string" || candidate.length > 2048) {
    console.error("Unexpected video response:", JSON.stringify(result)?.slice(0, 500));
    throw videoError("The provider finished but returned no usable video. Please try again.", 502, true);
  }

  let parsed;
  try {
    parsed = new URL(candidate);
  } catch {
    throw videoError("The provider returned an invalid video link. Please try again.", 502, true);
  }

  if (parsed.protocol !== "https:") {
    throw videoError("The provider returned an unsafe video link. Please try again.", 502, true);
  }

  return parsed.toString();
}

async function runVideoModel(env, { prompt, duration, aspectRatio }) {
  let timeoutId;

  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(
      () =>
        reject(
          videoError(
            "The video is taking longer than expected. It may still finish on the provider's side, so your seconds were not returned. Please wait a few minutes before trying again.",
            504,
            false
          )
        ),
      VIDEO_STUDIO.timeoutMs
    );
  });

  let result;

  try {
    result = await Promise.race([
      env.AI.run(VIDEO_STUDIO.model, {
        prompt,
        aspect_ratio: aspectRatio,
        duration,
        resolution: VIDEO_STUDIO.resolution
      }),
      timeout
    ]);
  } catch (cause) {
    if (cause && cause.refund === false) throw cause;

    const message = String(cause?.message || cause);
    console.error("Video model error:", message);

    if (/moderat|safety|policy|prohibit|not allowed|violat/i.test(message)) {
      throw videoError(
        "That prompt could not be used because it may break content guidelines. Please rephrase it and try again.",
        400,
        true
      );
    }
    if (/invalid|bad request|unprocessable|validation/i.test(message)) {
      throw videoError("The video settings or prompt were not accepted. Please adjust them and try again.", 400, true);
    }
    if (/quota|credit|balance|billing|insufficient|limit|429/i.test(message)) {
      throw videoError("Video generation is busy or temporarily unavailable. Please try again later.", 503, true);
    }
    throw videoError("The video provider could not complete the request. Please try again.", 502, true);
  } finally {
    clearTimeout(timeoutId);
  }

  return extractVideoUrl(result);
}

/* ------------------------------------------------------------------ */
/* Engineering Drawing Studio — one image in, all views out            */
/* ------------------------------------------------------------------ */

// The AI looks at ONE picture of a part (any view, drawing, sketch or photo),
// describes the part as simple 3D solids, and this code draws every view
// from that description as SVG.
//
// Coordinates: x = width (left to right in the front view),
//              y = height (up), z = depth (toward the viewer in the front view).


const DRAWING_PROMPT = [
  "You are an engineering part reconstruction system, not an illustration generator.",
  "Inspect the uploaded image carefully before creating the model. The image may show a detailed mechanical component, assembly, sketch, technical drawing, photograph, or isometric view.",
  "Reconstruct the specific object visible in the image. Do not replace it with a generic block, cylinder, or simplified object if additional features are visible.",
  "First identify the major body and EACH visibly distinct feature: stepped levels, bases, platforms, uprights, brackets, jaws, shafts, bosses, flanges, collars, holes, slots, recesses, grooves, ribs, handles, and cylindrical parts. Include only features supported by visible evidence.",
  "Represent a stepped or irregular body using multiple appropriately sized and positioned boxes rather than one oversized box. Represent visible round shafts, bosses, and holes with cylinders. Use op=cut only for holes, bores, or slots that are actually visible or strongly supported by the drawing. Do not invent hidden features.",
  "The renderer supports boxes and cylinders only. Use combinations of these primitives to approximate the actual visible geometry as closely as possible. Do not describe details in prose instead of encoding them in parts.",
  "Coordinates: x = width (left to right in the front view), y = height (up), z = depth. Keep all dimensions positive and place the complete object near x=0, y=0, z=0.",
  "Use the printed unit and dimensions exactly when legible. Otherwise estimate relative proportions from the image, set estimated=true, and do not pretend estimates are measured dimensions.",
  "Use these exact primitive formats:",
  '{"type":"box","op":"add","x":0,"y":0,"z":0,"w":40,"h":20,"d":30}',
  '{"type":"cyl","op":"add","axis":"x","cx":20,"cy":10,"cz":15,"r":5,"len":20}',
  "For cylinders, axis must be x, y, or z; cx, cy, cz are the cylinder centre coordinates; r is radius; len is length along the axis.",
  "For holes, bores, or slots use op=cut and place cuts after the add solids. Do not use cut to represent a protruding part.",
  "Use as many as 24 solids when needed to preserve visible shape details. Prefer a faithful multi-part reconstruction over a minimal primitive count. Avoid duplicate, irrelevant, or overlapping solids unless the physical object requires them.",
  "Before returning JSON, check that the parts collectively resemble the image, that the major components have not been omitted, and that dimensions and positions are internally consistent. If the image is too unclear to identify the shape, return an empty parts array and a short title indicating uncertainty rather than guessing a completely different object.",
  "Return ONLY valid JSON with no Markdown or commentary, using this exact structure:",
  '{"title":"short part name","unit":"mm","estimated":true,"parts":[{"type":"box","op":"add","x":0,"y":0,"z":0,"w":40,"h":20,"d":30}]}'
].join("\n");


const DRAW = {
  W: 900,
  H: 620,
  bg: "#0b0d10",
  line: "#d3ec8a",
  area: { x: 140, y: 70, w: 620, h: 390 }
};

const AXES = ["x", "y", "z"];
const rnd = n => Math.round(n * 100) / 100;
const fmtN = n => String(Math.round(n * 100) / 100);
const cleanXml = s => String(s).replace(/[<>&"']/g, "");

function clampNum(value, lo, hi, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.min(hi, Math.max(lo, n)) : fallback;
}

/* ---------- Check and clean what the AI returned ---------- */

function sanitizeModel(raw) {
  const list = Array.isArray(raw?.parts) ? raw.parts.slice(0, 24) : [];
  const adds = [];
  const cuts = [];

  for (const p of list) {
    const type = String(p?.type || "").toLowerCase();
    const op = String(p?.op || "add").toLowerCase() === "cut" ? "cut" : "add";
    let part = null;

    if (type === "box") {
      const w = clampNum(p.w, 0, 5000, 0);
      const h = clampNum(p.h, 0, 5000, 0);
      const d = clampNum(p.d, 0, 5000, 0);

      if (w > 0 && h > 0 && d > 0) {
        part = {
          type: "box",
          op,
          x: clampNum(p.x, -5000, 5000, 0),
          y: clampNum(p.y, -5000, 5000, 0),
          z: clampNum(p.z, -5000, 5000, 0),
          w,
          h,
          d
        };
      }
    } else if (type === "cyl" || type === "cylinder") {
      const axis = String(p.axis || "z").toLowerCase();
      const r = clampNum(
        p.r ?? Number(p.diameter ?? p.dia) / 2,
        0,
        2500,
        0
      );
      const len = clampNum(p.len ?? p.length, 0, 5000, 0);

      if (AXES.includes(axis) && r > 0 && len > 0) {
        part = {
          type: "cyl",
          op,
          axis,
          r,
          len,
          c: {
            x: clampNum(p.cx ?? p.x, -5000, 5000, 0),
            y: clampNum(p.cy ?? p.y, -5000, 5000, 0),
            z: clampNum(p.cz ?? p.z, -5000, 5000, 0)
          }
        };
      }
    }

    if (part) (op === "cut" ? cuts : adds).push(part);
  }

  if (!adds.length) {
    const error = new Error(
      "I could not work out the part's shape from that image. Please try a clearer picture, ideally with the dimensions printed on it."
    );
    error.status = 400;
    throw error;
  }

  return {
    title: cleanXml(String(raw?.title || "Machine part")).slice(0, 60),
    unit: String(raw?.unit || "").toLowerCase() === "in" ? "in" : "mm",
    estimated: Boolean(raw?.estimated),
    parts: [...adds, ...cuts]
  };
}

function extentsOf(p) {
  if (p.type === "box") {
    return {
      x: [p.x, p.x + p.w],
      y: [p.y, p.y + p.h],
      z: [p.z, p.z + p.d]
    };
  }

  const e = {};
  for (const a of AXES) {
    const half = a === p.axis ? p.len / 2 : p.r;
    e[a] = [p.c[a] - half, p.c[a] + half];
  }
  return e;
}

/* ---------- Flat views: front, top, right side ---------- */

const VIEWS = {
  front: { u: "x", v: "y", su: 1, sv: 1, depth: "z", label: "FRONT VIEW" },
  top: { u: "x", v: "z", su: 1, sv: -1, depth: "y", label: "TOP VIEW" },
  side: { u: "z", v: "y", su: -1, sv: 1, depth: "x", label: "RIGHT VIEW (FROM RIGHT)" }
};

function projectPart(p, V, depthRange) {
  const e = extentsOf(p);
  const spansDepth =
    e[V.depth][0] <= depthRange[0] + 1e-6 &&
    e[V.depth][1] >= depthRange[1] - 1e-6;

  const base = {
    op: p.op,
    through: spansDepth,
    cyl: p.type === "cyl",
    horiz: p.type === "cyl" && p.axis === V.u
  };

  if (p.type === "cyl" && p.axis === V.depth) {
    return {
      ...base,
      kind: "circ",
      cu: V.su * p.c[V.u],
      cv: V.sv * p.c[V.v],
      r: p.r
    };
  }

  const ua = [V.su * e[V.u][0], V.su * e[V.u][1]].sort((a, b) => a - b);
  const va = [V.sv * e[V.v][0], V.sv * e[V.v][1]].sort((a, b) => a - b);

  return { ...base, kind: "rect", u0: ua[0], u1: ua[1], v0: va[0], v1: va[1] };
}

// Cuts the part with a plane through its middle (z = zc) and returns the
// cut faces, seen from the front.
function sectionPart(p, zc) {
  const base = { op: p.op, through: true, cyl: false, horiz: false };

  if (p.type === "box") {
    if (zc < p.z || zc > p.z + p.d) return null;
    return {
      ...base,
      kind: "rect",
      u0: p.x,
      u1: p.x + p.w,
      v0: p.y,
      v1: p.y + p.h
    };
  }

  const dz = zc - p.c.z;

  if (p.axis === "z") {
    if (Math.abs(dz) > p.len / 2) return null;
    return { ...base, kind: "circ", cu: p.c.x, cv: p.c.y, r: p.r };
  }

  if (Math.abs(dz) >= p.r) return null;
  const hr = Math.sqrt(p.r * p.r - dz * dz);

  if (p.axis === "x") {
    return {
      ...base,
      kind: "rect",
      u0: p.c.x - p.len / 2,
      u1: p.c.x + p.len / 2,
      v0: p.c.y - hr,
      v1: p.c.y + hr
    };
  }

  return {
    ...base,
    kind: "rect",
    u0: p.c.x - hr,
    u1: p.c.x + hr,
    v0: p.c.y - p.len / 2,
    v1: p.c.y + p.len / 2
  };
}

function bounds2D(shapes) {
  let u0 = Infinity, u1 = -Infinity, v0 = Infinity, v1 = -Infinity;

  for (const s of shapes) {
    if (s.kind === "circ") {
      u0 = Math.min(u0, s.cu - s.r);
      u1 = Math.max(u1, s.cu + s.r);
      v0 = Math.min(v0, s.cv - s.r);
      v1 = Math.max(v1, s.cv + s.r);
    } else {
      u0 = Math.min(u0, s.u0);
      u1 = Math.max(u1, s.u1);
      v0 = Math.min(v0, s.v0);
      v1 = Math.max(v1, s.v1);
    }
  }

  return { u0, u1, v0, v1 };
}

function fitTo(b) {
  const a = DRAW.area;
  const bw = Math.max(b.u1 - b.u0, 1e-6);
  const bh = Math.max(b.v1 - b.v0, 1e-6);
  const s = Math.min(a.w / bw, a.h / bh);
  const ox = a.x + (a.w - bw * s) / 2;
  const oy = a.y + (a.h - bh * s) / 2;

  return {
    s,
    X: u => ox + (u - b.u0) * s,
    Y: v => oy + (b.v1 - v) * s
  };
}

/* ---------- SVG building blocks ---------- */

const svgLine = (x1, y1, x2, y2, extra = "") =>
  `<line x1="${rnd(x1)}" y1="${rnd(y1)}" x2="${rnd(x2)}" y2="${rnd(y2)}" stroke="${DRAW.line}" stroke-width="1" ${extra}/>`;

function svgArrow(x, y, dx, dy) {
  const a = 7;
  const b = 2.6;
  const px = -dy;
  const py = dx;

  return `<polygon points="${rnd(x)},${rnd(y)} ${rnd(x - dx * a + px * b)},${rnd(y - dy * a + py * b)} ${rnd(x - dx * a - px * b)},${rnd(y - dy * a - py * b)}" fill="${DRAW.line}"/>`;
}

function svgText(x, y, text, extra = "") {
  return `<text x="${rnd(x)}" y="${rnd(y)}" fill="${DRAW.line}" font-family="Arial, Helvetica, sans-serif" font-size="12" ${extra}>${cleanXml(text)}</text>`;
}

function hDim(x1, x2, y, label, refY) {
  return (
    svgLine(x1, refY, x1, y + 5) +
    svgLine(x2, refY, x2, y + 5) +
    svgLine(x1, y, x2, y) +
    svgArrow(x1, y, -1, 0) +
    svgArrow(x2, y, 1, 0) +
    svgText((x1 + x2) / 2, y - 5, label, 'text-anchor="middle"')
  );
}

function vDim(x, y1, y2, label, refX) {
  const mid = (y1 + y2) / 2;

  return (
    svgLine(refX, y1, x + 5, y1) +
    svgLine(refX, y2, x + 5, y2) +
    svgLine(x, y1, x, y2) +
    svgArrow(x, y1, 0, -1) +
    svgArrow(x, y2, 0, 1) +
    `<text transform="translate(${rnd(x + 14)},${rnd(mid)}) rotate(-90)" text-anchor="middle" fill="${DRAW.line}" font-family="Arial, Helvetica, sans-serif" font-size="12">${cleanXml(label)}</text>`
  );
}

function frameSvg(model, projection, label, body) {
  const pj = projection === "first" ? "FIRST-ANGLE" : "THIRD-ANGLE";
  const note =
    "AI-generated from image" +
    (model.estimated ? " (dimensions estimated)" : "") +
    " - verify before use";

  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${DRAW.W} ${DRAW.H}" width="100%" role="img" style="background:${DRAW.bg};border-radius:12px;display:block">` +
    `<defs><pattern id="hatch" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="8" stroke="${DRAW.line}" stroke-width="1"/></pattern></defs>` +
    `<rect x="14" y="14" width="${DRAW.W - 28}" height="${DRAW.H - 28}" fill="none" stroke="${DRAW.line}" stroke-width="1.2"/>` +
    body +
    svgText(DRAW.W / 2, 548, label, 'text-anchor="middle" font-weight="bold"') +
    svgLine(14, 572, DRAW.W - 14, 572) +
    svgText(24, 592, `${model.title} | Units: ${model.unit} | ${pj} PROJECTION | Fitted scale, not 1:1`, 'font-size="10"') +
    svgText(DRAW.W - 24, 592, note, 'font-size="10" text-anchor="end"') +
    `</svg>`
  );
}

// Draws a set of flat shapes (front/top/side/section) with outlines,
// holes, centre lines and dimensions.
function drawFlat(shapes, opts) {
  const adds = shapes.filter(s => s.op === "add");
  const cuts = shapes.filter(s => s.op === "cut");
  const b = bounds2D(adds);
  const f = fitTo(b);

  const el = (s, attrs) =>
    s.kind === "circ"
      ? `<circle cx="${rnd(f.X(s.cu))}" cy="${rnd(f.Y(s.cv))}" r="${rnd(s.r * f.s)}" ${attrs}/>`
      : `<rect x="${rnd(f.X(s.u0))}" y="${rnd(f.Y(s.v1))}" width="${rnd((s.u1 - s.u0) * f.s)}" height="${rnd((s.v1 - s.v0) * f.s)}" ${attrs}/>`;

  let out = "";

  // Thick outline of the combined shape (hides seams between joined solids)
  out += adds.map(s => el(s, `fill="${DRAW.line}" stroke="${DRAW.line}" stroke-width="3.2"`)).join("");
  out += adds.map(s => el(s, `fill="${DRAW.bg}"`)).join("");
  if (opts.hatch) out += adds.map(s => el(s, `fill="url(#hatch)"`)).join("");
  // Thin lines where separate solids meet
  out += adds.map(s => el(s, `fill="none" stroke="${DRAW.line}" stroke-width=".8" opacity=".5"`)).join("");

  // Holes: solid if they go all the way through, dashed (hidden) otherwise
  for (const s of cuts) {
    out += s.through
      ? el(s, `fill="${DRAW.bg}" stroke="${DRAW.line}" stroke-width="1.6"`)
      : el(s, `fill="none" stroke="${DRAW.line}" stroke-width="1" stroke-dasharray="5 3.5"`);
  }

  // Centre lines
  if (opts.centre) {
    const dash = 'stroke-dasharray="14 3 2 3"';

    for (const s of shapes) {
      if (s.kind === "circ") {
        const cx = f.X(s.cu);
        const cy = f.Y(s.cv);
        const R = s.r * f.s + 10;
        out += svgLine(cx - R, cy, cx + R, cy, dash) + svgLine(cx, cy - R, cx, cy + R, dash);
      } else if (s.cyl) {
        const x0 = f.X(s.u0);
        const x1 = f.X(s.u1);
        const y0 = f.Y(s.v1);
        const y1 = f.Y(s.v0);

        out += s.horiz
          ? svgLine(x0 - 8, (y0 + y1) / 2, x1 + 8, (y0 + y1) / 2, dash)
          : svgLine((x0 + x1) / 2, y0 - 8, (x0 + x1) / 2, y1 + 8, dash);
      }
    }
  }

  // Overall dimensions
  const X0 = f.X(b.u0);
  const X1 = f.X(b.u1);
  const Yt = f.Y(b.v1);
  const Yb = f.Y(b.v0);

  out += hDim(X0, X1, Yb + 30, fmtN(b.u1 - b.u0), Yb);
  out += vDim(X1 + 34, Yt, Yb, fmtN(b.v1 - b.v0), X1);

  // Diameters of visible circles
  const seen = new Set();
  const dirs = [[0.707, -0.707], [-0.707, -0.707], [0.707, 0.707], [-0.707, 0.707]];
  let n = 0;

  for (const s of shapes) {
    if (s.kind !== "circ" || n >= 4) continue;
    const key = fmtN(s.r);
    if (seen.has(key)) continue;
    seen.add(key);

    const [dx, dy] = dirs[n % 4];
    const cx = f.X(s.cu);
    const cy = f.Y(s.cv);
    const R = s.r * f.s;

    out += svgLine(cx + dx * R, cy + dy * R, cx + dx * (R + 22), cy + dy * (R + 22));
    out += svgText(
      cx + dx * (R + 26),
      cy + dy * (R + 26) + 4,
      "\u00D8" + fmtN(s.r * 2),
      `text-anchor="${dx > 0 ? "start" : "end"}"`
    );
    n++;
  }

  return out;
}

function flatViewSvg(model, key, projection) {
  const V = VIEWS[key];
  const adds3 = model.parts.filter(p => p.op === "add");
  const depthRange = [
    Math.min(...adds3.map(p => extentsOf(p)[V.depth][0])),
    Math.max(...adds3.map(p => extentsOf(p)[V.depth][1]))
  ];

  const shapes = model.parts.map(p => projectPart(p, V, depthRange));

  return frameSvg(
    model,
    projection,
    V.label,
    drawFlat(shapes, { hatch: false, centre: true })
  );
}

function sectionViewSvg(model, projection) {
  const adds3 = model.parts.filter(p => p.op === "add");
  const zc =
    (Math.min(...adds3.map(p => extentsOf(p).z[0])) +
      Math.max(...adds3.map(p => extentsOf(p).z[1]))) / 2;

  const shapes = model.parts.map(p => sectionPart(p, zc)).filter(Boolean);

  if (!shapes.some(s => s.op === "add")) {
    return flatViewSvg(model, "front", projection);
  }

  return frameSvg(
    model,
    projection,
    "SECTION A-A (FULL SECTION)",
    drawFlat(shapes, { hatch: true, centre: false })
  );
}

/* ---------- Isometric view ---------- */

// Viewer looks from the +x, +y, +z corner. x runs right-down, z runs left-down.
const isoPt = (x, y, z) => [(x - z) * 0.866025, y - (x + z) * 0.5];

function convexHull(points) {
  const pts = points
    .map(p => [p[0], p[1]])
    .sort((a, b) => a[0] - b[0] || a[1] - b[1]);

  if (pts.length < 3) return pts;

  const cross = (o, a, b) =>
    (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);

  const lower = [];
  for (const p of pts) {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], p) <= 0) lower.pop();
    lower.push(p);
  }

  const upper = [];
  for (let i = pts.length - 1; i >= 0; i--) {
    const p = pts[i];
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], p) <= 0) upper.pop();
    upper.push(p);
  }

  lower.pop();
  upper.pop();
  return lower.concat(upper);
}

function ringPoints(p, endSign) {
  const others = AXES.filter(a => a !== p.axis);
  const pts = [];
  const N = 40;

  for (let i = 0; i < N; i++) {
    const t = (2 * Math.PI * i) / N;
    const q = { x: 0, y: 0, z: 0 };
    q[p.axis] = p.c[p.axis] + (endSign * p.len) / 2;
    q[others[0]] = p.c[others[0]] + p.r * Math.cos(t);
    q[others[1]] = p.c[others[1]] + p.r * Math.sin(t);
    pts.push(isoPt(q.x, q.y, q.z));
  }

  return pts;
}

function isoGeometry(p) {
  if (p.type === "box") {
    const x0 = p.x, x1 = p.x + p.w;
    const y0 = p.y, y1 = p.y + p.h;
    const z0 = p.z, z1 = p.z + p.d;
    const P = (x, y, z) => isoPt(x, y, z);

    const corners = [
      P(x0, y0, z0), P(x1, y0, z0), P(x1, y0, z1), P(x0, y0, z1),
      P(x0, y1, z0), P(x1, y1, z0), P(x1, y1, z1), P(x0, y1, z1)
    ];

    return {
      hull: convexHull(corners),
      faces: [
        { pts: [P(x0, y1, z0), P(x1, y1, z0), P(x1, y1, z1), P(x0, y1, z1)], alpha: 0.24 },
        { pts: [P(x0, y0, z1), P(x1, y0, z1), P(x1, y1, z1), P(x0, y1, z1)], alpha: 0.1 },
        { pts: [P(x1, y0, z0), P(x1, y0, z1), P(x1, y1, z1), P(x1, y1, z0)], alpha: 0.17 }
      ],
      top: [P(x0, y1, z0), P(x1, y1, z0), P(x1, y1, z1), P(x0, y1, z1)]
    };
  }

  const near = ringPoints(p, 1);
  const far = ringPoints(p, -1);

  return {
    hull: convexHull(near.concat(far)),
    faces: [{ pts: near, alpha: 0.2 }],
    cap: near
  };
}

function isoViewSvg(model, projection) {
  const adds = model.parts
    .filter(p => p.op === "add")
    .sort((a, b) => depthKey(a) - depthKey(b));
  const cuts = model.parts.filter(p => p.op === "cut");

  function depthKey(p) {
    const e = extentsOf(p);
    return (e.x[0] + e.x[1] + e.y[0] + e.y[1] + e.z[0] + e.z[1]) / 2;
  }

  const addGeo = adds.map(isoGeometry);
  const cutGeo = cuts.map(isoGeometry);

  const all = addGeo.flatMap(g => g.hull);
  const b = {
    u0: Math.min(...all.map(p => p[0])),
    u1: Math.max(...all.map(p => p[0])),
    v0: Math.min(...all.map(p => p[1])),
    v1: Math.max(...all.map(p => p[1]))
  };
  const f = fitTo(b);

  const poly = (pts, attrs) =>
    `<polygon points="${pts.map(p => `${rnd(f.X(p[0]))},${rnd(f.Y(p[1]))}`).join(" ")}" ${attrs}/>`;

  let out = "";

  for (const g of addGeo) {
    out += poly(g.hull, `fill="${DRAW.bg}"`);
    for (const face of g.faces) {
      out += poly(face.pts, `fill="${DRAW.line}" fill-opacity="${face.alpha}" stroke="${DRAW.line}" stroke-width="1"`);
    }
    out += poly(g.hull, `fill="none" stroke="${DRAW.line}" stroke-width="1.5" stroke-linejoin="round"`);
  }

  for (const g of cutGeo) {
    const shape = g.cap || g.top;
    out += poly(shape, `fill="${DRAW.bg}" stroke="${DRAW.line}" stroke-width="1.4"`);
  }

  return frameSvg(model, projection, "ISOMETRIC VIEW (APPROXIMATE)", out);
}

/* ---------- All views at once ---------- */

function buildAllViews(model, projection) {
  return {
    front: flatViewSvg(model, "front", projection),
    top: flatViewSvg(model, "top", projection),
    side: flatViewSvg(model, "side", projection),
    iso: isoViewSvg(model, projection),
    section: sectionViewSvg(model, projection),
    aux: null
  };
}

/* ------------------------------------------------------------------ */
/* CAD code validation                                                 */
/* ------------------------------------------------------------------ */


function validateCadCode(rawCode) {
  let code = String(rawCode || "").trim();

  code = code
    .replace(/^```(?:javascript|js)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();

  if (!code || code.length > 20000) {
    throw new Error(
      "The generated CAD script was empty or too long. Please try again."
    );
  }

  // Scan the source while preserving string literals safely.
  // Comments are removed, and template strings are not supported.
  let masked = "";
  const strings = [];
  let i = 0;

  while (i < code.length) {
    const ch = code[i];
    const next = code[i + 1];

    if (ch === "`") {
      throw new Error("Template strings are not supported in CAD scripts.");
    }

    if (ch === "/" && next === "/") {
      i += 2;
      while (i < code.length && code[i] !== "\n") i++;
      masked += "\n";
      continue;
    }

    if (ch === "/" && next === "*") {
      i += 2;
      while (
        i < code.length &&
        !(code[i] === "*" && code[i + 1] === "/")
      ) i++;

      if (i >= code.length) {
        throw new Error("The generated CAD script has an unclosed comment.");
      }

      i += 2;
      masked += "  ";
      continue;
    }

    if (ch === "'" || ch === '"') {
      const quote = ch;
      let literal = "";
      i++;

      while (i < code.length && code[i] !== quote) {
        if (code[i] === "\\") {
          throw new Error(
            "Escaped characters are not supported in CAD text."
          );
        }

        if (code.charCodeAt(i) < 32) {
          throw new Error("Invalid character in a CAD text literal.");
        }

        literal += code[i++];
      }

      if (i >= code.length) {
        throw new Error("The generated CAD script has an unclosed string.");
      }

      i++;
      strings.push(literal);
      masked += " STRLITERAL ";
      continue;
    }

    masked += ch;
    i++;
  }

  // Never permit general-purpose JavaScript or browser access.
  const forbidden =
    /\b(?:window|document|globalThis|self|fetch|XMLHttpRequest|WebSocket|import|export|eval|Function|constructor|prototype|__proto__|localStorage|sessionStorage|navigator|location|process|require|Worker|setTimeout|setInterval|postMessage|cookie|indexedDB|await|async|class|while|for|do|try|catch|throw|debugger|this|delete|with|yield|function|new\s+Function|Proxy|Reflect|WebAssembly|SharedArrayBuffer|Atomics|arguments|super)\b/i;

  const forbiddenMatch = masked.match(forbidden);

if (forbiddenMatch) {
  throw new Error(
    "Unsupported JavaScript token: " +
    forbiddenMatch[0] +
    ". Please regenerate the CAD model."
  );
}

  // Only allow documented CAD operations and their supported methods.
  const allowedFunctions = new Set([
    "Box", "Sphere", "Cylinder", "Cone", "Circle", "Polygon",
    "BSpline", "Text3D", "Sketch", "Translate", "Rotate", "Scale",
    "Mirror", "Union", "Difference", "Intersection", "Extrude",
    "Revolve", "Loft", "Pipe", "Offset", "FilletEdges",
    "ChamferEdges", "Edges", "Volume", "SurfaceArea", "CenterOfMass"
  ]);

  const allowedMethods = new Set([
    "LineTo", "Fillet", "ArcTo", "BSplineTo", "End", "Face", "Wire",
    "ofType", "parallel", "max", "min", "indices"
  ]);

  const allowedKeywords = new Set([
    "const", "let", "var", "return", "new",
    "true", "false", "null", "undefined"
  ]);

  const declared = new Set();

  for (const match of masked.matchAll(
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)/g
  )) {
    declared.add(match[1]);
  }

  // Restrict property access to approved CAD methods.
  for (const match of masked.matchAll(
    /\.\s*([A-Za-z_$][\w$]*)/g
  )) {
    const property = match[1];

    if (
      !allowedMethods.has(property) &&
      !["PI", "abs", "min", "max", "sqrt", "sin", "cos", "tan",
        "round", "floor", "ceil", "pow"].includes(property)
    ) {
      throw new Error(
        "The CAD script uses an unsupported property: " + property
      );
    }
  }

  // Reject computed property access and indirect function calls.
  if (
    /\b[A-Za-z_$][\w$]*\s*\[\s*[A-Za-z_$]/.test(masked) ||
    /\]\s*\(/.test(masked) ||
    /\)\s*\(/.test(masked) ||
    /=>|\?\./.test(masked)
  ) {
    throw new Error(
      "The generated CAD script uses unsupported JavaScript syntax."
    );
  }

  // Reject object literals, function bodies, and other block syntax.
  if (/[{}]/.test(masked)) {
    throw new Error(
      "Use CAD operations and variable declarations only."
    );
  }

  const identifiers = masked.match(/[A-Za-z_$][\w$]*/g) || [];

  for (const identifier of identifiers) {
    if (
      identifier === "STRLITERAL" ||
      allowedKeywords.has(identifier) ||
      allowedFunctions.has(identifier) ||
      allowedMethods.has(identifier) ||
      declared.has(identifier)
    ) {
      continue;
    }

    // Math is allowed only with the approved properties listed above.
    if (identifier === "Math") continue;

    throw new Error(
      "The CAD script uses an unsupported identifier: " + identifier
    );
  }

  // Permit only safe literal text for plane and selector names.
  for (const literal of strings) {
    if (
      literal.length > 500 ||
      /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(literal)
    ) {
      throw new Error("The CAD script contains invalid text.");
    }
  }

  if (
    ![...allowedFunctions].some(name =>
      new RegExp("\\b" + name + "\\s*\\(").test(masked)
    )
  ) {
    throw new Error(
      "The CAD script did not contain a supported geometry operation."
    );
  }

  // Check bracket balance.
  let depth = 0;
  const pairs = { ")": "(", "]": "[" };
  const stack = [];

  for (const ch of masked) {
    if (ch === "(" || ch === "[") {
      stack.push(ch);
      depth++;
    } else if (ch === ")" || ch === "]") {
      if (stack.pop() !== pairs[ch]) {
        throw new Error("The generated CAD script has invalid brackets.");
      }
      depth--;
    }

    if (depth < 0) {
      throw new Error("The generated CAD script has invalid brackets.");
    }
  }

  if (depth !== 0) {
    throw new Error("The generated CAD script has invalid brackets.");
  }

  return code;
}

/* ------------------------------------------------------------------ */
/* Worker                                                              */
/* ------------------------------------------------------------------ */

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    const url = new URL(request.url);
    const path = url.pathname;

    /* ---------------- Health ---------------- */
    if (request.method === "GET" && path === "/api/health") {
      return respond({
        ok: true,
        service: "Malie AI",
        provider: "Cloudflare Workers AI",
        endpoints: [
          "/api/health",
          "/api/chat",
          "/api/write",
          "/api/quiz",
          "/api/excel",
          "/api/subscription/status",
          "/api/payment/request",
          "/api/payment/submit-reference",
          "/api/drawing",
          "/api/drawing",
          "/api/cad/generate",
          "/api/video/access",
          "/api/video/generate"
        ]
      });
    }

    /* ---------------- Subscription status ---------------- */
    if (request.method === "POST" && path === "/api/subscription/status") {
      try {
        const userId = await authenticateFirebase(request, env);

        if (!env.DB) {
          return respond(
            { error: "Subscription service is temporarily unavailable." },
            503
          );
        }

        const subscription = await env.DB.prepare(
          `SELECT plan, status, current_period_end
           FROM subscriptions
           WHERE user_id = ?
           LIMIT 1`
        ).bind(userId).first();

        const expiry = subscription?.current_period_end
          ? Date.parse(subscription.current_period_end)
          : null;

        const active =
          String(subscription?.status || "").toLowerCase() === "active" &&
          expiry !== null &&
          Number.isFinite(expiry) &&
          expiry > Date.now();

        const latest = await env.DB.prepare(
          `SELECT id, plan, amount, billing_period, status, payment_reference
           FROM payment_requests
           WHERE user_id = ?
           ORDER BY id DESC
           LIMIT 1`
        ).bind(userId).first();

        return respond({
          ok: true,
          subscription: {
            active,
            plan: active
              ? String(subscription.plan || "").toLowerCase()
              : "free",
            expiresAt: active ? subscription.current_period_end : null
          },
          latestRequest: latest
            ? {
                requestId: latest.id,
                plan: latest.plan,
                billingPeriod: latest.billing_period,
                amount: latest.amount,
                status: latest.status,
                hasReference: Boolean(latest.payment_reference)
              }
            : null
        });
      } catch (error) {
        console.error("Subscription status error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        return respond({ error: "Could not load your plan." }, 500);
      }
    }

    /* ---------------- Create payment request ---------------- */
    if (request.method === "POST" && path === "/api/payment/request") {
      try {
        const userId = await authenticateFirebase(request, env);

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const plan = String(body.plan || "").toLowerCase();
        const billingPeriod = String(
          body.billingPeriod || "monthly"
        ).toLowerCase();

        const amount = PRICES[plan]?.[billingPeriod];

        if (!amount) {
          return respond({ error: "Invalid plan or billing period." }, 400);
        }

        if (!env.DB) {
          return respond(
            { error: "Payment service is temporarily unavailable." },
            503
          );
        }

        const created = await env.DB.prepare(
          `INSERT INTO payment_requests
             (user_id, plan, amount, billing_period, status)
           VALUES (?, ?, ?, ?, 'pending')
           RETURNING id`
        ).bind(userId, plan, amount, billingPeriod).first();

        if (!created?.id) {
          throw new Error("Payment request could not be created.");
        }

        return respond({
          ok: true,
          requestId: created.id,
          plan,
          billingPeriod,
          amount,
          currency: "NGN",
          paymentDetails: {
            bankName: "FirstBank",
            accountName: env.FIRSTBANK_ACCOUNT_NAME || null,
            accountNumber: env.FIRSTBANK_ACCOUNT_NUMBER || null
          },
          message:
            "Payment request created. Make the transfer, then submit your transfer reference."
        });
      } catch (error) {
        console.error("Payment request error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        return respond(
          { error: "Could not create your payment request." },
          500
        );
      }
    }

    /* ---------------- Submit transfer reference ---------------- */
    if (request.method === "POST" && path === "/api/payment/submit-reference") {
      try {
        const userId = await authenticateFirebase(request, env);

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const requestId = Number(body.requestId);
        const transferReference = String(body.transferReference || "").trim();

        if (
          !Number.isSafeInteger(requestId) ||
          requestId < 1 ||
          transferReference.length < 4 ||
          transferReference.length > 120
        ) {
          return respond(
            { error: "Please provide a valid request ID and transfer reference." },
            400
          );
        }

        if (!env.DB) {
          return respond(
            { error: "Payment service is temporarily unavailable." },
            503
          );
        }

        const result = await env.DB.prepare(
          `UPDATE payment_requests
           SET payment_reference = ?
           WHERE id = ?
             AND user_id = ?
             AND status = 'pending'`
        ).bind(transferReference, requestId, userId).run();

        if (!result.success) {
          throw new Error("Reference could not be saved.");
        }

        if (result.meta?.changes !== 1) {
          return respond(
            {
              error:
                "Payment request not found, already processed, or reference already used."
            },
            409
          );
        }

        return respond({
          ok: true,
          message:
            "Transfer reference submitted. Your payment must be verified before your subscription is activated."
        });
      } catch (error) {
        console.error("Payment reference error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        if (String(error.message || "").includes("UNIQUE constraint failed")) {
          return respond(
            { error: "That transfer reference has already been submitted." },
            409
          );
        }

        return respond(
          { error: "Could not submit your transfer reference." },
          500
        );
      }
    }

    /* ---------------- Admin: pending payments ---------------- */
    if (request.method === "POST" && path === "/api/admin/pending-payments") {
      try {
        const adminUid = env.MALIE_ADMIN_UID;

        if (!adminUid) {
          return respond({ error: "Admin access is not configured." }, 500);
        }

        const authenticatedUid = await authenticateFirebase(request, env);

        if (authenticatedUid !== adminUid) {
          return respond({ error: "Unauthorized." }, 403);
        }

        if (!env.DB) {
          return respond({ error: "Payment service unavailable." }, 503);
        }

        const result = await env.DB.prepare(
          `SELECT id, user_id, plan, amount, billing_period,
                  payment_reference, status, created_at
           FROM payment_requests
           WHERE status = 'pending'
           ORDER BY created_at ASC
           LIMIT 100`
        ).all();

        return respond({
          ok: true,
          payments: (result.results || []).map(payment => ({
            requestId: payment.id,
            userId: payment.user_id,
            plan: payment.plan,
            amount: payment.amount,
            billingPeriod: payment.billing_period,
            hasReference: Boolean(payment.payment_reference),
            transferReference: payment.payment_reference,
            status: payment.status,
            createdAt: payment.created_at
          }))
        });
      } catch (error) {
        console.error("Admin pending-payments error:", error.message);

        return respond(
          { error: error.message || "Could not load pending payments." },
          error.status || 500
        );
      }
    }

    /* ---------------- Admin: approve payment ---------------- */
    if (request.method === "POST" && path === "/api/admin/approve-payment") {
      try {
        const adminUid = env.MALIE_ADMIN_UID;

        if (!adminUid) {
          return respond({ error: "Admin access is not configured." }, 500);
        }

        let authenticatedUid;

        try {
          authenticatedUid = await authenticateFirebase(request, env);
        } catch (error) {
          return respond(
            { error: error.message || "Unauthorized." },
            error.status || 401
          );
        }

        if (authenticatedUid !== adminUid) {
          return respond({ error: "Unauthorized." }, 403);
        }

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const requestId = Number(body.requestId);

        if (!Number.isSafeInteger(requestId) || requestId < 1) {
          return respond({ error: "Invalid payment request ID." }, 400);
        }

        if (!env.DB) {
          return respond({ error: "Payment service unavailable." }, 503);
        }

        const payment = await env.DB.prepare(
          `SELECT id, user_id, plan, amount, billing_period,
                  payment_reference, status
           FROM payment_requests
           WHERE id = ?`
        ).bind(requestId).first();

        if (!payment) {
          return respond({ error: "Payment request not found." }, 404);
        }

        if (payment.status !== "pending") {
          return respond(
            { error: "This payment request has already been processed." },
            409
          );
        }

        if (!payment.payment_reference) {
          return respond(
            { error: "The customer has not submitted a transfer reference." },
            400
          );
        }

        if (
          !PRICES[payment.plan] ||
          PRICES[payment.plan][payment.billing_period] !== payment.amount
        ) {
          return respond(
            { error: "Payment amount or plan details do not match." },
            400
          );
        }

        // Verify the actual bank transfer before calling this route.
        const now = new Date();
        const confirmedAt = now.toISOString();
        const endDate = new Date(now);

        if (payment.billing_period === "annual") {
          endDate.setUTCFullYear(endDate.getUTCFullYear() + 1);
        } else {
          endDate.setUTCMonth(endDate.getUTCMonth() + 1);
        }

        const results = await env.DB.batch([
          env.DB.prepare(
            `UPDATE payment_requests
             SET status = 'confirmed', confirmed_at = ?
             WHERE id = ?
               AND status = 'pending'
               AND payment_reference IS NOT NULL`
          ).bind(confirmedAt, requestId),

          env.DB.prepare(
            `INSERT INTO subscriptions
               (user_id, plan, status, current_period_end,
                payment_reference, updated_at)
             SELECT user_id, plan, 'active', ?, payment_reference, ?
             FROM payment_requests
             WHERE id = ?
               AND status = 'confirmed'
               AND confirmed_at = ?
             ON CONFLICT(user_id) DO UPDATE SET
               plan = excluded.plan,
               status = 'active',
               current_period_end = excluded.current_period_end,
               payment_reference = excluded.payment_reference,
               updated_at = excluded.updated_at`
          ).bind(endDate.toISOString(), confirmedAt, requestId, confirmedAt)
        ]);

        if (!results[0]?.meta?.changes) {
          return respond(
            { error: "Payment request could not be confirmed. Check its status." },
            409
          );
        }

        return respond({
          ok: true,
          message: "Payment confirmed and subscription activated.",
          requestId,
          plan: payment.plan,
          billingPeriod: payment.billing_period,
          expiresAt: endDate.toISOString()
        });
      } catch (error) {
        console.error("Admin payment approval error:", error.message);
        return respond({ error: "Could not approve this payment." }, 500);
      }
    }

    /* ---------------- Chat / Write / Quiz / Excel ---------------- */
    if (
      request.method === "POST" &&
      ["/api/chat", "/api/write", "/api/quiz", "/api/excel"].includes(path)
    ) {
      let body;

      try {
        body = await request.json();
      } catch {
        return respond({ error: "Please send valid JSON." }, 400);
      }

      try {
        const userId = await authenticateFirebase(request, env);

        /* ----- Chat ----- */
        if (path === "/api/chat") {
          const message = String(body.message || "").trim();

          if (!message) {
            return respond({ error: "Please enter a message." }, 400);
          }

          if (message.length > 30000) {
            return respond(
              { error: "Your message is too long. Please shorten it." },
              400
            );
          }

          const requestedMode = String(body.mode || "").trim().toLowerCase();

          // Look up the plan once; reuse it for the Excel check and usage limit.
          const plan = await getActivePlan(env, userId);

          if (requestedMode === "excel") {
            // Excel Assistant is available to Premium and Pro subscribers.
            if (plan !== "premium" && plan !== "pro") {
              return respond(
                { error: "Please subscribe to Premium or Pro to use this feature." },
                403
              );
            }
          }

          const usage = await reserveDailyUsage(env, userId, plan);

          const identityQuestion =
            /\b(?:who\s+(?:created|made|built|developed|owns|founded)\s+(?:you|malie(?:\s+ai)?|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?(?:the\s+)?(?:creator|founder|owner)\s+of\s+(?:malie(?:\s+ai)?|you|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?your\s+(?:creator|founder|owner)|who\s+is\s+behind\s+(?:malie(?:\s+ai)?|you)|(?:are\s+you|is\s+malie(?:\s+ai)?)\s+(?:created|made|built|developed)\s+by)\b/i;

          if (identityQuestion.test(message)) {
            return respond({
              reply:
                "I am Malie AI, developed by Campus Creative, founded by Osume Prosper. Google provides the model technology; Campus Creative created the Malie AI product.",
              usage
            });
          }

          let reply;

          try {
            reply = await callAI(
              env,
              [
                "Respond to the user's message as Malie AI.",
                "If the user is learning something, teach it in clear steps.",
                "For calculations, show the working and explain how the result was obtained.",
                "For school subjects, use simple explanations and examples appropriate to the question.",
                "If the user asks for a direct answer, give it first and then explain it briefly.",

                ...(requestedMode === "excel"
                  ? [
                      "You are now Malie AI's Premium Excel Assistant.",
                      "Help with Excel formulas, spreadsheet assignments, data analysis, and troubleshooting spreadsheet errors.",
                      "For formulas, give the exact formula, explain where to enter it, and explain how it works.",
                      "For calculations, show the steps and check the arithmetic where possible.",
                      "For spreadsheet errors, explain the likely cause and provide a specific fix.",
                      "When users paste tables, analyze only the data they provide. Never invent missing data.",
                      "If cell references or column meanings are unclear, ask a brief clarifying question.",
                      "If the user wants a downloadable .xlsx file, tell them to use the Excel File Builder, where they can paste a table or upload a photo and download the finished spreadsheet."
                    ]
                  : []),

                "",
                "User message:",
                message
              ].join("\n")
            );
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          return respond({ reply, usage });
        }

        /* ----- Write ----- */
        if (path === "/api/write") {
          const mode = String(body.mode || "correct").trim();
          const inputText = String(body.text || "").trim();
          const options = body.options || {};

          if (!inputText) {
            return respond({ error: "Please enter some text." }, 400);
          }

          if (inputText.length > 30000) {
            return respond(
              { error: "Your text is too long. Please shorten it and try again." },
              400
            );
          }

          const instruction =
            taskInstructions[mode] || taskInstructions.correct;

          const usage = await reserveDailyUsage(env, userId);

          let result;

          try {
            result = await callAI(
              env,
              [
                "Task: " + instruction,
                "User options: " + JSON.stringify(options),
                "Follow the selected task carefully.",
                "Preserve names, dates, figures, and other supplied facts.",
                "Do not add unsupported information.",
                "",
                "Text to process:",
                inputText
              ].join("\n")
            );
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          return respond({ result, usage });
        }

        /* ----- Quiz ----- */
        if (path === "/api/quiz") {
          const inputText = String(body.text || "").trim();

          const count = Math.max(
            1,
            Math.min(20, Number.parseInt(body.count, 10) || 5)
          );

          const difficulty = String(body.difficulty || "medium").trim();

          if (!inputText) {
            return respond(
              { error: "Please provide notes or text to generate a quiz." },
              400
            );
          }

          if (inputText.length > 30000) {
            return respond(
              { error: "Your source text is too long. Please shorten it and try again." },
              400
            );
          }

          const prompt = [
            "Create exactly " + count + " multiple-choice questions based ONLY on the source material below.",
            "Difficulty: " + difficulty + ".",
            "Each question must have exactly four answer options.",
            "There must be exactly one correct answer per question.",
            "Vary the correct answer position across questions.",
            "Avoid duplicate questions and unsupported facts.",
            "Make the questions test understanding, not only memorisation, when appropriate.",
            "The explanation must teach the learner why the correct option is right.",
            "For calculations, include the key reasoning or working in the explanation.",
            "Keep explanations concise, clear, and suitable for the requested difficulty.",
            "Use answerIndex as a zero-based number: 0 for the first option, 1 for the second, 2 for the third, or 3 for the fourth.",
            "Do not put letters such as A or B in answerIndex.",
            "Return ONLY valid JSON in this exact structure:",
            '{"questions":[{"question":"Question text","options":["Option 1","Option 2","Option 3","Option 4"],"answerIndex":0,"explanation":"Explain why the correct option is right."}]}',
            "",
            "Source material:",
            inputText
          ].join("\n");

          const resultText = await callAI(env, prompt, true);
          const quiz = parseQuiz(resultText);

          if (quiz.questions.length !== count) {
            throw new Error(
              `The AI returned ${quiz.questions.length} questions instead of ${count}. Please generate the quiz again.`
            );
          }

          return respond(quiz);
        }

        /* ----- Excel file builder (Premium / Pro) ----- */
        if (path === "/api/excel") {
          const plan = await getActivePlan(env, userId);

          if (plan !== "premium" && plan !== "pro") {
            return respond(
              { error: "Please subscribe to Premium or Pro to use this feature." },
              403
            );
          }

          const message = String(body.message || "").trim();
          const image = body.image ? String(body.image) : null;

          if (!message && !image) {
            return respond(
              { error: "Send a table, a photo, or instructions." },
              400
            );
          }

          if (message.length > 30000) {
            return respond({ error: "Your message is too long." }, 400);
          }

          if (
            image &&
            (image.length > 5_000_000 ||
              !/^data:image\/(png|jpeg|webp);base64,/.test(image))
          ) {
            return respond(
              { error: "Please upload a PNG, JPG or WebP image under about 3.5 MB." },
              400
            );
          }

          const usage = await reserveDailyUsage(env, userId, plan);

          const prompt = [
            "Convert the user's table/figures (text and/or image) into a spreadsheet.",
            "Copy values exactly. Never invent data. If something is unreadable, leave the cell empty.",
            "Put headers in row 1. Use real numbers (not strings) for numeric values.",
            "Where the user asks for totals, averages or calculations, use Excel formulas as strings starting with '=' (for example \"=SUM(B2:B10)\").",
            "Return ONLY valid JSON in this exact structure:",
            '{"filename":"short_name","sheets":[{"name":"Sheet1","rows":[["Header A","Header B"],["x",1]]}]}',
            "",
            "User instructions/data:",
            message || "(see image)"
          ].join("\n");

          let fileBase64;
          let spec;

          try {
            const text = await callAI(env, prompt, true, image);
            spec = parseJsonObject(
              text,
              "The AI did not return a valid spreadsheet. Please try again."
            );
            fileBase64 = buildWorkbook(spec);
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          const filename =
            (String(spec.filename || "malie_sheet")
              .replace(/[^\w-]/g, "_")
              .slice(0, 40) || "malie_sheet") + ".xlsx";

          return respond({
            ok: true,
            filename,
            fileBase64,
            preview: spec.sheets?.[0]?.rows?.slice(0, 10) || [],
            usage
          });
        }
      } catch (error) {
        const status = getErrorStatus(error);

        console.error("API error:", { path, status, message: error.message });

        return respond(
          { error: error.message || "The request failed. Please try again." },
          status
        );
      }
    }

    /* ---------------- Admin check (any signed-in user) ---------------- */
    if (request.method === "POST" && path === "/api/admin/check") {
      try {
        const uid = await authenticateFirebase(request, env);
        const adminUid = env.MALIE_ADMIN_UID;
        return respond({ ok: true, isAdmin: Boolean(adminUid) && uid === adminUid });
      } catch (error) {
        return respond({ ok: false, isAdmin: false, error: error.message }, error.status || 401);
      }
    }

    /* Advert Studio access check — does not start a render */
    if (request.method === "POST" && path === "/api/advert/access") {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond({
            error: "Malie AI Advert Studio is available on the Pro plan only."
          }, 403);
        }

        return respond({ ok: true, plan: "pro" });
      } catch (e) {
        return respond({
          error: e.message || "Could not verify Advert Studio access."
        }, e.status || 500);
      }
    }

    /* ---------------- Malie AI Advert Studio (Pro only) ---------------- */
    if (
      request.method === "POST" &&
      (path === "/api/advert/start" || path === "/api/advert/status")
    ) {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond({
            error: "Malie AI Advert Studio is available on the Pro plan only."
          }, 403);
        }

        if (!env.ETERNAL_AI_API_KEY) {
          return respond({ error: "Advert Studio is not configured yet." }, 500);
        }

        if (!env.DB) {
          return respond({ error: "The database is not configured." }, 500);
        }

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const apiHeaders = {
          "Authorization": `Bearer ${env.ETERNAL_AI_API_KEY}`,
          "Content-Type": "application/json"
        };

        if (path === "/api/advert/status") {
          const projectId = String(body.projectId || "");

          if (!/^[A-Za-z0-9-]{16,80}$/.test(projectId)) {
            return respond({ error: "Invalid video project ID." }, 400);
          }

          const owned = await env.DB.prepare(
            "SELECT project_id, status, video_url FROM advert_projects WHERE project_id = ? AND user_id = ?"
          ).bind(projectId, userId).first();

          if (!owned) {
            return respond({ error: "Video project not found for this account." }, 404);
          }

          if (owned.status === "completed" && owned.video_url) {
            return respond({
              ok: true,
              projectId,
              status: "completed",
              videoUrl: owned.video_url,
              message: "Your advert video is ready. Save it promptly; the provider URL may expire after about 24 hours."
            });
          }

          const statusResponse = await fetch(
            "https://open.eternalai.org/api/image-to-video/" +
              encodeURIComponent(projectId) + "/status",
            {
              headers: { "Authorization": `Bearer ${env.ETERNAL_AI_API_KEY}` },
              signal: AbortSignal.timeout(15000)
            }
          );

          const statusData = await statusResponse.json().catch(() => ({}));
          const result = statusData.result || {};
          console.log("ETERNAL STATUS HTTP:", statusResponse.status);
          console.log("ETERNAL STATUS BODY:", JSON.stringify(statusData));

          if (!statusResponse.ok || (!statusData.result && statusData.status !== true)) {
            return respond({
              error: statusData.error || "Could not retrieve the video status. Please try again."
            }, 502);
          }

          const providerStatus = String(result.status || "processing").toLowerCase();

          // Eternal AI puts the finished video at result.output.video_url
          // (other shapes are checked too).
          const foundUrl =
            result.output?.video_url ||
            result.video_url ||
            (Array.isArray(result.output)
              ? result.output[0]?.video_url || result.output[0]
              : null) ||
            result.output?.url ||
            result.url ||
            null;

          const DONE = ["success", "succeeded", "completed", "complete", "ready", "done", "finished"];
          const FAILED = ["failed", "error", "cancelled", "canceled", "timeout"];

          let status;
          if (typeof foundUrl === "string" && foundUrl) status = "completed";
          else if (DONE.includes(providerStatus)) status = "completed";
          else if (FAILED.includes(providerStatus)) status = "failed";
          else status = providerStatus; // pending / processing / queued

          const videoUrl =
            status === "completed"
              ? (typeof foundUrl === "string" ? foundUrl : null)
              : owned.video_url;

          await env.DB.prepare(
            "UPDATE advert_projects SET status = ?, video_url = ? WHERE project_id = ? AND user_id = ?"
          ).bind(status, videoUrl || null, projectId, userId).run();

          return respond({
            ok: true,
            projectId,
            status,
            videoUrl: videoUrl || null,
            progress: result.progress ?? null,
            message: status === "completed"
              ? "Your advert video is ready. Save it promptly; the provider URL may expire after about 24 hours."
              : status === "failed"
                ? (result.error || statusData.error || "The video generation did not complete.")
                : "Your advert video is still rendering."
          });
        }

        const today = getUsageDate();
        const usage = await env.DB.prepare(
          "SELECT COUNT(*) AS total FROM advert_projects WHERE user_id = ? AND usage_date = ?"
        ).bind(userId, today).first();

        if (Number(usage?.total || 0) >= 3) {
          return respond({
            error: "You have reached the limit of 3 advert video renders per day. Please try again tomorrow."
          }, 429);
        }

        const imageData = String(body.imageData || "");
        const match = imageData.match(
          /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=\r\n]+)$/
        );

        if (!match) {
          return respond({ error: "Upload a valid JPG, PNG, or WebP product image." }, 400);
        }

        const base64 = match[2].replace(/\s/g, "");

        if (!base64 || base64.length > 20000000) {
          return respond({ error: "Image is too large. Please use an image smaller than 15 MB." }, 413);
        }

        let imageBytes;
        try {
          imageBytes = atob(base64);
        } catch {
          return respond({ error: "The image data is invalid." }, 400);
        }

        if (!imageBytes.length || imageBytes.length > 15 * 1024 * 1024) {
          return respond({ error: "Image must be smaller than 15 MB." }, 413);
        }

        const cleanText = (value, maxLength) => String(value || "")
          .replace(/[<>]/g, "")
          .trim()
          .slice(0, maxLength);

        let rawHeadline = cleanText(body.headline, 400);
        let rawDescription = cleanText(body.description, 400);
        const split = rawHeadline.match(/^\s*(?:headline\s*[:\-]\s*)?(.*?)\s*description\s*[:\-]\s*(.*)$/i);

        if (split) {
          rawHeadline = split[1];
          if (!rawDescription) rawDescription = split[2];
        }

        const stripLabel = (value, label) =>
          value.replace(new RegExp("^\\s*" + label + "\\s*[:\\-]\\s*", "i"), "").trim();

        const headline = stripLabel(rawHeadline, "headline").slice(0, 70) || "A Product Worth Discovering";
        const description = stripLabel(rawDescription, "description").slice(0, 120);
        const callToAction = stripLabel(cleanText(body.callToAction, 40), "call to action") || "Shop Now";

        const videoStyle = ["clean", "fullscreen", "promo", "story"].includes(body.style)
          ? body.style
          : "clean";

        const styleDirection = {
          clean: "Elegant, clean studio commercial. Keep the complete product visible and centred with generous negative space; use a slow cinematic push-in and soft premium lighting.",
          fullscreen: "Full-frame cinematic product commercial with tasteful camera motion and dramatic but realistic lighting. Keep the product recognizable and do not crop important product details.",
          promo: "Energetic premium promotional commercial with a smooth camera move, polished lighting changes, and a refined product reveal.",
          story: "Vertical social-story product commercial, composed for a 9:16 mobile frame. Keep the product fully visible and away from the edges, with smooth cinematic motion."
        }[videoStyle];

        const promptParts = [
          "Create a polished 5-second product advertisement from the supplied image.",
          styleDirection,
          "Preserve the product's exact shape, proportions, colours, packaging, label and existing logo. Do not redesign, replace, or distort the product.",
          "Use realistic studio lighting, natural reflections, subtle atmospheric movement, and smooth camera motion. Professional commercial quality.",
          `Headline concept: ${headline}.`,
          description ? `Supporting message: ${description}.` : "",
          `Call to action concept: ${callToAction}.`,
          "Do not generate new lettering, captions, fake labels, watermarks, or logos; keep any existing product text unchanged."
        ].filter(Boolean).join(" ").slice(0, 2000);

        const renderResponse = await fetch("https://open.eternalai.org/api/image-to-video", {
          method: "POST",
          headers: apiHeaders,
          body: JSON.stringify({
            prompt: promptParts,
            image_url: imageData,
            model_id: "wan-ai/wan2.2-i2v-a14b-lightning",
            duration: "5",
            resolution: "480p",
            aspect_ratio: videoStyle === "story" ? "9:16" : "16:9",
            negative_prompt: "blurry, distorted product, warped packaging, changed label, extra objects, illegible text, low quality, watermark",
            cfg_scale: 0.5
          }),
          signal: AbortSignal.timeout(30000)
        });

        const renderData = await renderResponse.json().catch(() => ({}));
        const requestId = renderData.result?.request_id;

        if (!renderResponse.ok || !requestId) {
          return respond({
            error: renderData.error || "Eternal AI could not start the video generation. Please try again.",
            provider_status: renderResponse.status
          }, renderResponse.status === 402 ? 402 : 502);
        }

        await env.DB.prepare(
          "INSERT INTO advert_projects (project_id, user_id, usage_date, status) VALUES (?, ?, ?, ?)"
        ).bind(requestId, userId, today, "pending").run();

        return respond({
          ok: true,
          projectId: requestId,
          status: "pending",
          message: "Your advert video has been submitted to Eternal AI. It may take a few minutes to finish.",
          durationSeconds: 5,
          resolution: "480p"
        });

      } catch (error) {
        console.error("Advert Studio error:", error);
        return respond({
          error: "Advert Studio encountered an error. Please try again."
        }, 500);
      }
    }
    /* ---------------- CAD Studio code generation (Pro only) ---------------- */

if (request.method === "POST" && path === "/api/cad/generate") {
  let body;

  try {
    body = await request.json();
  } catch {
    return respond({ error: "Please send valid JSON." }, 400);
  }

  try {
    const userId = await authenticateFirebase(request, env);
    const plan = await getActivePlan(env, userId);

    if (plan !== "pro") {
      return respond({
        error: "AI CAD generation is available on the Pro plan. Please upgrade to Pro to use it."
      }, 403);
    }

    const description = String(
      body.prompt || body.description || ""
    ).trim();

    if (description.length < 5) {
      return respond({
        error: "Describe the 3D object you want to create."
      }, 400);
    }

    if (description.length > 2000) {
      return respond({
        error: "Please keep your CAD description under 2,000 characters."
      }, 400);
    }

    if (!env.DB) {
      return respond({
        error: "Usage tracking is temporarily unavailable."
      }, 503);
    }

    const usage = await reserveDailyUsage(env, userId, plan);

    try {
      



const prompt = [
  "Generate reliable 3D CAD JavaScript for Cascade Studio.",
  "Return ONLY a JSON object with exactly two fields: title (short string) and code (string). No markdown fences.",
  "The code must use only the documented CascadeStudio standard library.",
  "Supported primitives: Box(x,y,z), Sphere(r), Cylinder(r,h), Cone(r1,r2,h), Circle(r,wire?), Polygon(points,wire?), Text3D(text,size,height,font?), BSpline(points,closed?).",
  "Supported sketch API: new Sketch([x,y], plane?).LineTo([x,y]).Fillet(r).ArcTo([midX,midY],[endX,endY]).BSplineTo(points).End(true).Face(). The plane may be XY, XZ, or YZ.",
  "Supported operations: Translate([x,y,z],shape), Rotate([ax,ay,az],degrees,shape), Scale(factor,shape), Mirror([x,y,z],shape), Union(shapes), Difference(body,[tools]), Intersection(shapes), Extrude(face,[dx,dy,dz]), Revolve(shape,degrees), Loft(wires), Pipe(shape,path), Offset(shape,distance), FilletEdges(shape,radius,indices), ChamferEdges(shape,distance,indices).",
  "For Pipe(shape, wirePath, keepInputs), the second argument MUST be a valid TopoDS_Wire, not a Sketch object or a general shape.",
"For hollow pipes, create the outer and inner pipes using the SAME centerline wire and pass true as the third Pipe argument to preserve the shared inputs.",
"Subtract the inner pipe from the outer pipe using Difference(outerPipe, [innerPipe]) and assign the result to finalModel.",
"Never use return finalModel; at the end of the generated script.",
  "Supported selectors and measurements: Edges(shape).ofType(type).parallel(axis).max(axis).min(axis).indices(), Volume(shape), SurfaceArea(shape), CenterOfMass(shape).",
  "For square or rectangular profiles, prefer Sketch with 2D points followed by End(true).Face() and Extrude(face,[0,0,height]).",
  "Example rectangular block: const face = new Sketch([-20,-20]).LineTo([20,-20]).LineTo([20,20]).LineTo([-20,20]).End(true).Face(); const finalModel = Extrude(face,[0,0,20]);",
  "Polygon requires an array of 3D points. Correct example: const face = Polygon([[0,0,0],[40,0,0],[40,40,0],[0,40,0]]); const finalModel = Extrude(face,[0,0,20]); Never pass a flat list of coordinates to Polygon.",
  "For a hollow cylinder, use Difference(Cylinder(outerRadius,height),[Translate([0,0,-1],Cylinder(innerRadius,height+2))]). Ensure the inner radius is smaller than the outer radius.",
  "For holes, use cylindrical cutters that extend at least 1 mm beyond both faces. Subtract the cutters from the actual body using Difference.",
  "Assign the completed requested model to a variable named finalModel. Preserve the result of every Boolean operation.",

  "Use const or let declarations and numeric coordinates. Use positive dimensions and sensible proportions. Negative coordinates are allowed in point arrays and translations.",
  "Do not use imports, custom functions, loops, conditions, browser APIs, arbitrary JavaScript, or markdown fences. Do not invent unsupported CAD functions.",
  
  "CRITICAL CODE RESTRICTIONS:",
  "Never use for, while, or do loops.",
  "Never use map, forEach, filter, or reduce.",
  "Do not generate helper functions or callbacks.",
  "Create geometry using direct, explicit calls to the documented CascadeStudio API.",
  "For repeated features such as bolt holes, write each operation explicitly.",
  "Return only valid JSON containing the model title and code.",
  "Before returning, check that the code contains no forbidden JavaScript keywords or unsupported syntax.",
  "Generate actual 3D geometry, not a description. Choose operations appropriate to the requested shape and its dimensions.",
  "User requested object:",
  description
].join("\n");




     
      const raw = await callAI(env, prompt, true);

      const result = parseJsonObject(
        raw,
        "Malie could not generate CAD code from that description. Please try again."
      );

      let title = String(
        result.title || "Generated CAD Model"
      ).replace(/[<>&"']/g, "").slice(0, 80);

      let code;

      try {
        code = validateCadCode(result.code);
      } catch (validationError) {
        const repairPrompt = [
          "Repair the invalid CascadeStudio CAD JavaScript below.",
          "The validation error was: " + String(validationError.message || validationError),
          "Return ONLY valid JSON with exactly two fields: title and code.",
          "Preserve the requested object's shape and dimensions.",
          "Use only documented CascadeStudio geometry operations.",
          "Do not use loops, helper functions, comments, template strings, or unsupported JavaScript.",
          "Ensure every parenthesis and square bracket is correctly balanced and nested.",
          "Assign the completed model to finalModel.",
          "Requested object:",
          description,
          "Previous title:",
          title,
          "Previous invalid code:",
          String(result.code || "")
        ].join("\n");

        const repairedRaw = await callAI(env, repairPrompt, true);

        const repairedResult = parseJsonObject(
          repairedRaw,
          "Malie could not repair the CAD script. Please generate the model again."
        );

        title = String(
          repairedResult.title || title
        ).replace(/[<>&"']/g, "").slice(0, 80);

        code = validateCadCode(repairedResult.code);
      }
 

      return respond({
        ok: true,
        title,
        code,
        usage
      });
    } catch (error) {
      await refundDailyUsage(env, userId, usage.date);
      throw error;
    }
  } catch (error) {
    const status = [400, 401, 403, 429, 500, 503].includes(error?.status)
      ? error.status
      : 502;

    console.error(
      "CAD generation error:",
      String(error?.message || error)
    );

    return respond({
      error: error?.status
        ? error.message
        : (error?.message || "CAD generation failed. Please try again.")
    }, status);
  }
    }

    /* ---------------- Engineering Drawing Studio (Pro only) ---------------- */
    if (request.method === "POST" && path === "/api/drawing") {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond(
            { error: "The Engineering Drawing Studio is available on the Pro plan. Please upgrade to Pro to use it." },
            403
          );
        }

        let body;
        try { body = await request.json(); }
        catch { return respond({ error: "Please send valid JSON." }, 400); }

        /* ----- NEW: upload one image, get every view back ----- */
        if (body.action === "image") {
          const image = String(body.image || "");

          if (
            image.length > 5_000_000 ||
            !/^data:image\/(png|jpeg|webp);base64,/.test(image)
          ) {
            return respond(
              { error: "Please upload a PNG, JPG or WebP image under about 3.5 MB." },
              400
            );
          }

          const projection = body.projection === "first" ? "first" : "third";
          const usage = await reserveDailyUsage(env, userId, plan);

          try {
            const text = await callAI(env, DRAWING_PROMPT, true, image);
            const model = sanitizeModel(
              parseJsonObject(
                text,
                "The AI could not describe that part. Please try a clearer image."
              )
            );
            const views = buildAllViews(model, projection);

            return respond({
              ok: true,
              supported: true,
              mode: "image",
              title: model.title,
              unit: model.unit,
              estimated: model.estimated,
              projection,
              views,
              auxMessage:
                "An auxiliary view is only needed for inclined surfaces, so none is generated.",
              model,
              usage
            });
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }
        }

        /* ----- Original: draw the stepped flanged bush from typed dimensions ----- */
        const raw = body.params || {};
        const p = {
          d: Number(raw.d),
          D1: Number(raw.D1),
          L1: Number(raw.L1),
          D2: Number(raw.D2),
          T: Number(raw.T)
        };

        for (const [key, value] of Object.entries(p)) {
          if (!Number.isFinite(value) || value <= 0 || value > 1000) {
            return respond({ error: `Dimension "${key}" must be a number greater than 0 and not more than 1000.` }, 400);
          }
        }
        if (!(p.d < p.D1)) return respond({ error: "The bore diameter must be smaller than the body diameter." }, 400);
        if (!(p.D1 < p.D2)) return respond({ error: "The flange diameter must be larger than the body diameter." }, 400);

        const views = ["front", "top", "side", "iso", "section", "aux"];
        const view = views.includes(body.view) ? body.view : "section";
        const projection = body.projection === "first" ? "first" : "third";
        const unit = raw.unit === "in" ? "in" : "mm";
        const title = String(raw.title || "Stepped flanged bush").replace(/[<>&"']/g, "").slice(0, 60);

        if (view === "aux") {
          return respond({
            ok: true,
            supported: false,
            message: "An auxiliary view is only needed for inclined surfaces. This part is a body of revolution with no inclined faces, so no auxiliary view is generated."
          });
        }

        return respond({
          ok: true,
          supported: true,
          view,
          projection,
          params: { ...p, unit, title },
          totalLength: p.L1 + p.T
        });
      } catch (error) {
        console.error("Drawing error:", String(error?.message || error));
        return respond({ error: error.message || "Could not generate the drawing." }, getErrorStatus(error));
      }
    }

    /* ---------------- Malie Video Studio (Pro only) ---------------- */
    if (
      request.method === "POST" &&
      (path === "/api/video/access" || path === "/api/video/generate")
    ) {
      try {
        // 1. Who is this? (verified with Firebase, never trusted from the browser)
        const userId = await authenticateFirebase(request, env);

        // 2. Active Pro plan, read from the subscriptions table
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond(
            { error: "Malie Video Studio is available on the Pro plan only." },
            403
          );
        }

        const limit = getVideoDailyLimit(env);
        const secondsLimit = getVideoDailySeconds(env);
        const usageDate = getUsageDate();

        /* ----- Access check + remaining allowance (no generation, no cost) ----- */
        if (path === "/api/video/access") {
          const used = await getVideoUsage(env, userId, usageDate);

          return respond({
            ok: true,
            plan: "pro",
            usage: videoUsageSummary(limit, used.videos, used.seconds, secondsLimit, usageDate),
            options: {
              durations: VIDEO_STUDIO.durations.filter(
                seconds => seconds <= secondsLimit
              ),
              aspectRatios: VIDEO_STUDIO.aspectRatios,
              resolution: VIDEO_STUDIO.resolution,
              promptMin: VIDEO_STUDIO.promptMin,
              promptMax: VIDEO_STUDIO.promptMax
            }
          });
        }

        /* ----- Generate ----- */
        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const prompt = String(body.prompt || "").trim();
        const duration = Number(body.duration);
        const aspectRatio = String(body.aspectRatio || "");

        if (prompt.length < VIDEO_STUDIO.promptMin) {
          return respond(
            { error: `Please describe your video in at least ${VIDEO_STUDIO.promptMin} characters.` },
            400
          );
        }

        if (prompt.length > VIDEO_STUDIO.promptMax) {
          return respond(
            { error: `Please keep your description under ${VIDEO_STUDIO.promptMax} characters.` },
            400
          );
        }

        if (!VIDEO_STUDIO.durations.includes(duration)) {
          return respond({ error: "Choose a duration of 5 or 10 seconds." }, 400);
        }

        if (!VIDEO_STUDIO.aspectRatios.includes(aspectRatio)) {
          return respond({ error: "Choose Landscape (16:9), Portrait (9:16) or Square (1:1)." }, 400);
        }

        // Resolution is fixed on the server (480p) for cost control; any value
        // sent by the browser is ignored.

        if (!env.AI || typeof env.AI.run !== "function") {
          return respond(
            { error: "Video generation is not configured yet. The site owner must add the Workers AI binding named AI." },
            500
          );
        }

        // 3. Reserve one of today's videos BEFORE spending money. Atomic, so
        //    parallel requests cannot exceed the daily allowance.
        const usedAfterReserve = await reserveVideo(env, userId, usageDate, duration, limit, secondsLimit);

        let videoUrl;

        try {
          videoUrl = await runVideoModel(env, { prompt, duration, aspectRatio });
        } catch (error) {
          // 4. Give the video back when the failure is definite.
          if (error.refund !== false) {
            await refundVideo(env, userId, usageDate, duration);
          }
          throw error;
        }

        return respond({
          ok: true,
          videoUrl,
          durationSeconds: duration,
          aspectRatio,
          resolution: VIDEO_STUDIO.resolution,
          usage: videoUsageSummary(limit, usedAfterReserve.videos, usedAfterReserve.seconds, secondsLimit, usageDate)
        });
      } catch (error) {
        console.error("Video Studio error:", String(error?.message || error));

        // Errors we raised on purpose carry a status and a safe message.
        // Anything else (for example a missing table) gets a generic reply.
        const status = [400, 401, 403, 429, 500, 502, 503, 504].includes(error?.status)
          ? error.status
          : 500;

        return respond(
          {
            error: error?.status
              ? error.message
              : "Video Studio could not process your request. Please try again."
          },
          status
        );
      }
    }

    /* ---------------- Fallbacks ---------------- */
    if (path.startsWith("/api/")) {
      return respond({ error: "API endpoint not found." }, 404);
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Malie AI is running.", { headers: corsHeaders });
  }
};
