const taskInstructions = {
  correct:
    "Correct grammar, spelling, punctuation, and sentence errors. Preserve the writer's meaning, facts, and voice. Return the corrected text first, then a short list of important changes only if helpful. Do not invent facts or citations.",

  improve:
    "Improve clarity, flow, structure, word choice, and readability while preserving the writer's intent and key facts. Return the revised text, followed by concise notes on the main improvements. Do not invent facts or citations.",

  generate:
    "Create an original, useful first draft based on the user's instructions. Respect the requested audience and length. Organize it logically and clearly. Do not fabricate factual claims or sources.",

  explain:
    "Act as a supportive writing tutor. Identify grammar, punctuation, word-choice, clarity, and structure issues. Quote a short excerpt, explain the problem, and show a better version. Distinguish firm errors from optional style improvements. End with a corrected version when practical.",

  summarize:
    "Summarize the supplied text at the requested length. Preserve the main ideas and important qualifications. Do not add facts absent from the source.",

  tone:
    "Rewrite the supplied text in the requested tone for the intended audience. Preserve its meaning and key details. Do not add unsupported claims. Return the rewritten version only unless a brief note is needed."
};

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store"
    }
  });

const corsHeaders = {
  "access-control-allow-origin": "*",
  "access-control-allow-methods": "GET, POST, OPTIONS",
  "access-control-allow-headers": "Content-Type"
};

function respond(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders,
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store"
    }
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    // Allow browser preflight requests.
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders
      });
    }

    // Health check. Does not expose the secret.
    if (path === "/api/health" && request.method === "GET") {
      return respond({
        ok: true,
        aiConfigured: Boolean(env.GEMINI_API_KEY),
        endpoints: ["/api/chat", "/api/write"]
      });
    }

    const isChat = path === "/api/chat";
    const isWriting = path === "/api/write";

    if (!isChat && !isWriting) {
      // Serve website assets when configured.
      if (env.ASSETS && typeof env.ASSETS.fetch === "function") {
        return env.ASSETS.fetch(request);
      }

      return respond({
        error: "Route not found.",
        path
      }, 404);
    }

    if (request.method !== "POST") {
      return respond({
        error: "Use POST for this endpoint."
      }, 405);
    }

    if (!env.GEMINI_API_KEY) {
      return respond({
        error: "GEMINI_API_KEY is missing from Cloudflare Worker secrets."
      }, 503);
    }

    let body;

    try {
      body = await request.json();
    } catch {
      return respond({
        error: "Invalid JSON. Please send a valid request."
      }, 400);
    }

    // Chat sends { message: "..." }.
    // Writing tools send { mode, text, options }.
    const mode = isChat ? "generate" : body?.mode;
    const text = isChat ? body?.message : body?.text;
    const options = body?.options || {};

    if (!Object.prototype.hasOwnProperty.call(taskInstructions, mode)) {
      return respond({
        error: "Choose a valid writing mode."
      }, 400);
    }

    if (typeof text !== "string" || !text.trim()) {
      return respond({
        error: isChat
          ? "Please enter a message."
          : "Enter text or instructions first."
      }, 400);
    }

    if (text.length > 30000) {
      return respond({
        error: "Please keep each request under 30,000 characters."
      }, 413);
    }

    const tone =
      typeof options.tone === "string"
        ? options.tone.slice(0, 80)
        : "Professional";

    const length =
      typeof options.length === "string"
        ? options.length.slice(0, 40)
        : "Balanced";

    const audience =
      typeof options.audience === "string"
        ? options.audience.slice(0, 100)
        : "General audience";

    const systemPrompt = isChat
      ? "You are Oracle AI, a helpful general-purpose AI assistant. Answer questions clearly and accurately. Help with learning, writing, coding, brainstorming, and everyday tasks. Admit uncertainty when necessary. Do not invent facts or sources."
      : "You are Oracle, a careful and helpful AI writing assistant for students and professionals. Be clear, constructive, and honest. Preserve user intent. Never claim you verified sources or facts when you have not.\n\nTask-specific instructions: " +
        taskInstructions[mode];

    const userInput = isChat
      ? text.trim()
      : `Requested tone: ${tone}
Requested length: ${length}
Audience: ${audience}

User content and instructions:
${text}`;

    // Keep the model configurable in Cloudflare.
    const model = env.GEMINI_MODEL || "gemini-2.5-flash";

    const apiUrl =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      encodeURIComponent(model) +
      ":generateContent";

    try {
      const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-goog-api-key": env.GEMINI_API_KEY
        },
        body: JSON.stringify({
          systemInstruction: {
            parts: [{ text: systemPrompt }]
          },
          contents: [
            {
              role: "user",
              parts: [{ text: userInput }]
            }
          ],
          generationConfig: {
            maxOutputTokens: 3500,
            temperature: 0.7
          }
        })
      });

      const raw = await response.text();
      let data;

      try {
        data = raw ? JSON.parse(raw) : {};
      } catch {
        console.error("Gemini returned a non-JSON response.");
        return respond({
          error: "Gemini returned an unexpected response. Please try again."
        }, 502);
      }

      if (!response.ok) {
        // Log diagnostic details in Cloudflare, never log the API key.
        console.error("Gemini API error:", {
          status: response.status,
          message: data?.error?.message || "No error message returned",
          statusName: data?.error?.status || "UNKNOWN"
        });

        if (response.status === 429) {
          return respond({
            error: "Gemini quota or rate limit reached. Check your Google AI usage and billing, then try again."
          }, 429);
        }

        if (
          response.status === 400 ||
          response.status === 401 ||
          response.status === 403
        ) {
          return respond({
            error:
              "Gemini rejected the request. Check the model name, API key permissions, and Google AI API access. HTTP " +
              response.status
          }, 502);
        }

        if (response.status === 404) {
          return respond({
            error:
              "The configured Gemini model was not found. Check GEMINI_MODEL in Cloudflare. Current model: " +
              model
          }, 502);
        }

        return respond({
          error:
            "Gemini returned an error (HTTP " +
            response.status +
            "). Check the Worker logs for details."
        }, 502);
      }

      const parts = data?.candidates?.[0]?.content?.parts;

      const result = Array.isArray(parts)
        ? parts
            .map(part =>
              typeof part.text === "string" ? part.text : ""
            )
            .filter(Boolean)
            .join("")
            .trim()
        : "";

      if (!result) {
        console.error("Gemini response contained no text.", {
          finishReason: data?.candidates?.[0]?.finishReason || "UNKNOWN",
          promptFeedback: data?.promptFeedback || null
        });

        return respond({
          error:
            "Gemini returned no text. The request may have been blocked or the model may have reached a response limit."
        }, 502);
      }

      // Match the response expected by each frontend.
      if (isChat) {
        return respond({ reply: result });
      }

      return respond({ result });
    } catch (error) {
      console.error("Oracle AI fetch error:", {
        message: error?.message || "Unknown network error"
      });

      return respond({
        error:
          "Oracle could not reach Gemini. Check the Cloudflare Worker logs and try again."
      }, 502);
    }
  }
};
