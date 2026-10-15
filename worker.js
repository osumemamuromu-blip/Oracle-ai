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

// Daily request limits per plan. Change the numbers here only.
const DAILY_LIMITS = {
  free: 8,
  premium: 100,
  pro: 300
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization"
};

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

function parseQuiz(text) {
  const cleaned = text.trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "");

  let parsed;

  try {
    parsed = JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");

    if (start < 0 || end <= start) {
      throw new Error("The AI did not return valid quiz JSON. Please try again.");
    }

    parsed = JSON.parse(cleaned.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed.questions;

  if (!Array.isArray(questions) || !questions.length) {
    throw new Error("No quiz questions were returned. Please try again.");
  }

  const validated = questions.map((item, index) => {
    const question = String(
      item.question || item.prompt || ""
    ).trim();

    const options = item.options || item.choices;

    if (
      !question ||
      !Array.isArray(options) ||
      options.length < 2
    ) {
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
      explanation: String(
        item.explanation || item.reason || ""
      ).trim()
    };
  });

  return { questions: validated };
}

async function callAI(env, prompt, jsonMode = false) {
  if (!env.AI || typeof env.AI.run !== "function") {
    const error = new Error(
      "Cloudflare Workers AI is not connected. Check that your Worker has an AI binding named AI."
    );
    error.status = 500;
    throw error;
  }

  const model = "@cf/google/gemma-4-26b-a4b-it";

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

  let result;

  try {
    result = await env.AI.run(model, {
      messages: [
        {
          role: "system",
          content: systemPrompt
        },
        {
          role: "user",
          content: prompt
        }
      ],
      max_tokens: jsonMode ? 4096 : 2048,
      ...(jsonMode
        ? { response_format: { type: "json_object" } }
        : {}),
      chat_template_kwargs: {
        enable_thinking: false
      }
    });
  } catch (cause) {
    console.error("Workers AI error:", String(cause));

    const message = String(cause?.message || cause);

    const error = new Error(
      /limit|quota|neuron|429/i.test(message)
        ? "The Cloudflare AI usage limit may have been reached. Please try again after the usage limit resets."
        : /capacity|busy|overload|503/i.test(message)
        ? "Cloudflare AI is busy right now. Please wait a little and try again."
        : "Cloudflare AI could not complete the request. Please try again."
    );

    error.status =
      /limit|quota|neuron|429/i.test(message) ? 429 : 503;

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

function getErrorStatus(error) {
  if ([400, 401, 403, 429, 503].includes(error.status)) {
    return error.status;
  }

  return 500;
}

async function authenticateFirebase(request, env) {
  if (!env.FIREBASE_WEB_API_KEY) {
    const error = new Error(
      "Firebase verification is not configured."
    );
    error.status = 500;
    throw error;
  }

  const authorization = request.headers.get("Authorization") || "";
  const match = authorization.match(/^Bearer\s+(.+)$/i);

  if (!match) {
    const error = new Error(
      "Please sign in with Google to use Malie AI."
    );
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
    const error = new Error(
      "Could not verify your sign-in. Please try again."
    );
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

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders
      });
    }

    const url = new URL(request.url);
    const path = url.pathname;

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
          "/api/subscription/status",
          "/api/payment/request",
          "/api/payment/submit-reference"
        ]
      });
    }

    // Report the signed-in user's plan and latest payment request.
    if (
      request.method === "POST" &&
      path === "/api/subscription/status"
    ) {
      try {
        const userId = await authenticateFirebase(request, env);

        if (!env.DB) {
          return respond({
            error: "Subscription service is temporarily unavailable."
          }, 503);
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
          `SELECT id, plan, amount, billing_period, status,
                  payment_reference
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

        return respond({
          error: "Could not load your plan."
        }, 500);
      }
    }

    // Create a payment request for Premium or Pro.
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

        const prices = {
          premium: { monthly: 2500, annual: 25000 },
          pro: { monthly: 7500, annual: 85000 }
        };

        const amount = prices[plan]?.[billingPeriod];

        if (!amount) {
          return respond({
            error: "Invalid plan or billing period."
          }, 400);
        }

        if (!env.DB) {
          return respond({
            error: "Payment service is temporarily unavailable."
          }, 503);
        }

        const created = await env.DB.prepare(`
          INSERT INTO payment_requests
            (user_id, plan, amount, billing_period, status)
          VALUES (?, ?, ?, ?, 'pending')
          RETURNING id
        `).bind(userId, plan, amount, billingPeriod).first();

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

        return respond({
          error: "Could not create your payment request."
        }, 500);
      }
    }

    // Save a transfer reference without activating a subscription.
    if (
      request.method === "POST" &&
      path === "/api/payment/submit-reference"
    ) {
      try {
        const userId = await authenticateFirebase(request, env);
        const body = await request.json();

        const requestId = Number(body.requestId);
        const transferReference = String(
          body.transferReference || ""
        ).trim();

        if (
          !Number.isSafeInteger(requestId) ||
          requestId < 1 ||
          transferReference.length < 4 ||
          transferReference.length > 120
        ) {
          return respond({
            error: "Please provide a valid request ID and transfer reference."
          }, 400);
        }

        if (!env.DB) {
          return respond({
            error: "Payment service is temporarily unavailable."
          }, 503);
        }

        const result = await env.DB.prepare(`
          UPDATE payment_requests
          SET payment_reference = ?
          WHERE id = ?
            AND user_id = ?
            AND status = 'pending'
        `).bind(
          transferReference,
          requestId,
          userId
        ).run();

        if (!result.success) {
          throw new Error("Reference could not be saved.");
        }

        if (result.meta?.changes !== 1) {
          return respond({
            error: "Payment request not found, already processed, or reference already used."
          }, 409);
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

        if (
          String(error.message || "").includes("UNIQUE constraint failed")
        ) {
          return respond({
            error: "That transfer reference has already been submitted."
          }, 409);
        }

        return respond({
          error: "Could not submit your transfer reference."
        }, 500);
      }
    }

    
    // Admin-only list of pending payment requests.
    if (
      request.method === "POST" &&
      path === "/api/admin/pending-payments"
    ) {
      try {
        const adminUid = env.MALIE_ADMIN_UID;

        if (!adminUid) {
          return respond(
            { error: "Admin access is not configured." },
            500
          );
        }

        const authenticatedUid = await authenticateFirebase(request, env);

        if (authenticatedUid !== adminUid) {
          return respond({ error: "Unauthorized." }, 403);
        }

        if (!env.DB) {
          return respond(
            { error: "Payment service unavailable." },
            503
          );
        }

        const result = await env.DB.prepare(`
          SELECT id, user_id, plan, amount, billing_period,
                 payment_reference, status, created_at
          FROM payment_requests
          WHERE status = 'pending'
          ORDER BY created_at ASC
          LIMIT 100
        `).all();

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
          {
            error: error.message || "Could not load pending payments."
          },
          error.status || 500
        );
      }
    } 
    
    // Admin-only payment approval.
    if (
      request.method === "POST" &&
      path === "/api/admin/approve-payment"
    ) {
      try {
       const adminUid = env.MALIE_ADMIN_UID;

if (!adminUid) {
  return respond(
    { error: "Admin access is not configured." },
    500
  );
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
        
        const body = await request.json();
        const requestId = Number(body.requestId);

        if (!Number.isSafeInteger(requestId) || requestId < 1) {
          return respond({ error: "Invalid payment request ID." }, 400);
        }

        if (!env.DB) {
          return respond({ error: "Payment service unavailable." }, 503);
        }

        const payment = await env.DB.prepare(`
          SELECT id, user_id, plan, amount, billing_period,
                 payment_reference, status
          FROM payment_requests
          WHERE id = ?
        `).bind(requestId).first();

        if (!payment) {
          return respond({ error: "Payment request not found." }, 404);
        }

        if (payment.status !== "pending") {
          return respond({ error: "This payment request has already been processed." }, 409);
        }

        if (!payment.payment_reference) {
          return respond({ error: "The customer has not submitted a transfer reference." }, 400);
        }

        const expectedAmounts = {
          premium: { monthly: 2500, annual: 25000 },
          pro: { monthly: 7500, annual: 85000 }
        };

        if (
          !expectedAmounts[payment.plan] ||
          expectedAmounts[payment.plan][payment.billing_period] !== payment.amount
        ) {
          return respond({ error: "Payment amount or plan details do not match." }, 400);
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
          env.DB.prepare(`
            UPDATE payment_requests
            SET status = 'confirmed', confirmed_at = ?
            WHERE id = ?
              AND status = 'pending'
              AND payment_reference IS NOT NULL
          `).bind(confirmedAt, requestId),

          env.DB.prepare(`
            INSERT INTO subscriptions
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
              updated_at = excluded.updated_at
          `).bind(endDate.toISOString(), confirmedAt, requestId, confirmedAt)
        ]);

        if (!results[0]?.meta?.changes) {
          return respond({ error: "Payment request could not be confirmed. Check its status." }, 409);
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

    if (
      request.method === "POST" &&
      ["/api/chat", "/api/write", "/api/quiz"].includes(path)
    ) {
      let body;

      try {
        body = await request.json();
      } catch {
        return respond({ error: "Please send valid JSON." }, 400);
      }

      try {
        const userId = await authenticateFirebase(request, env);

        if (path === "/api/chat") {
          const message = String(body.message || "").trim();

          if (!message) {
            return respond({
              error: "Please enter a message."
            }, 400);
          }

          if (message.length > 30000) {
            return respond({
              error: "Your message is too long. Please shorten it."
            }, 400);
          }

          const requestedMode = String(body.mode || "").trim().toLowerCase();

          // Look up the plan once; reuse it for the Excel check and usage limit.
          const plan = await getActivePlan(env, userId);

          if (requestedMode === "excel") {
            // Excel Assistant is available to Premium and Pro subscribers.
            if (plan !== "premium" && plan !== "pro") {
              return respond({
                error: "Please subscribe to Premium or Pro to use this feature."
              }, 403);
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

          const reply = await callAI(
            env,
            [
              "Respond to the user's message as Malie AI.",
              "If the user is learning something, teach it in clear steps.",
              "For calculations, show the working and explain how the result was obtained.",
              "For school subjects, use simple explanations and examples appropriate to the question.",
              "If the user asks for a direct answer, give it first and then explain it briefly.",

              ...(requestedMode === "excel" ? [
                "You are now Malie AI's Premium Excel Assistant.",
                "Help with Excel formulas, spreadsheet assignments, data analysis, and troubleshooting spreadsheet errors.",
                "For formulas, give the exact formula, explain where to enter it, and explain how it works.",
                "For calculations, show the steps and check the arithmetic where possible.",
                "For spreadsheet errors, explain the likely cause and provide a specific fix.",
                "When users paste tables, analyze only the data they provide. Never invent missing data.",
                "If cell references or column meanings are unclear, ask a brief clarifying question.",
                "Do not claim to open, edit, or generate an Excel file unless a file-processing feature actually performed that action.",
                "Workbook uploads and downloadable Excel file editing are not available yet."
              ] : []),

              "",
              "User message:",
              message
            ].join("\n")
          );

          return respond({ reply, usage });
        }

        if (path === "/api/write") {
          const mode = String(body.mode || "correct").trim();
          const inputText = String(body.text || "").trim();
          const options = body.options || {};

          if (!inputText) {
            return respond({
              error: "Please enter some text."
            }, 400);
          }

          if (inputText.length > 30000) {
            return respond({
              error: "Your text is too long. Please shorten it and try again."
            }, 400);
          }

          const instruction =
            taskInstructions[mode] || taskInstructions.correct;

          const usage = await reserveDailyUsage(env, userId);

          const result = await callAI(
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

          return respond({ result, usage });
        }

        if (path === "/api/quiz") {
          const inputText = String(body.text || "").trim();

          const count = Math.max(
            1,
            Math.min(20, Number.parseInt(body.count, 10) || 5)
          );

          const difficulty = String(
            body.difficulty || "medium"
          ).trim();

          if (!inputText) {
            return respond({
              error: "Please provide notes or text to generate a quiz."
            }, 400);
          }

          if (inputText.length > 30000) {
            return respond({
              error: "Your source text is too long. Please shorten it and try again."
            }, 400);
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
      } catch (error) {
        const status = getErrorStatus(error);

        console.error("API error:", {
          path,
          status,
          message: error.message
        });

        return respond({
          error: error.message || "The request failed. Please try again."
        }, status);
      }
    }

    if (path.startsWith("/api/")) {
      return respond({
        error: "API endpoint not found."
      }, 404);
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Malie AI is running.", {
      headers: corsHeaders
    });
  }
};
