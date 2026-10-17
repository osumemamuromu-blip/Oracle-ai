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
  pro: { monthly: 7500, annual: 85000 }
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
          "/api/payment/submit-reference"
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

          /* ---- FIXED: find the video URL and recognise every finished/failed status ---- */
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
        return respond({ error: error.message || "Could not generate the drawing." }, getErrorStatus(error));
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
