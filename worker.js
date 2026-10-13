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

const sleep = ms =>
  new Promise(resolve => setTimeout(resolve, ms));

const transientStatuses = new Set([500, 502, 503, 504]);

export default {
  async fetch(request, env) {
    try {
      const startedAt = Date.now();
      const url = new URL(request.url);
      const path = url.pathname;

      if (request.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: corsHeaders
        });
      }

      if (path === "/api/health" && request.method === "GET") {
        return respond({
          ok: true,
          aiConfigured: Boolean(env.GEMINI_API_KEY),
          model: env.GEMINI_MODEL || "gemini-2.5-flash",
          endpoints: ["/api/chat", "/api/write"]
        });
      }

      const isChat = path === "/api/chat";
      const isWriting = path === "/api/write";

      if (!isChat && !isWriting) {
        if (
          env.ASSETS &&
          typeof env.ASSETS.fetch === "function"
        ) {
          return env.ASSETS.fetch(request);
        }

        return respond(
          { error: "Route not found." },
          404
        );
      }

      if (request.method !== "POST") {
        return respond(
          { error: "Use POST for this endpoint." },
          405
        );
      }

      if (!env.GEMINI_API_KEY) {
        console.error("GEMINI_API_KEY is missing.");

        return respond({
          error: "The AI service is not configured."
        }, 503);
      }

      let body;

      try {
        body = await request.json();
      } catch {
        return respond({
          error: "Invalid JSON request."
        }, 400);
      }

      const mode = isChat ? "generate" : body?.mode;
      const text = isChat ? body?.message : body?.text;
      const options = body?.options || {};

      if (
        !Object.prototype.hasOwnProperty.call(
          taskInstructions,
          mode
        )
      ) {
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
          error:
            "Please keep each request under 30,000 characters."
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
        ? "You are Malie AI, a helpful general-purpose AI assistant. Answer questions clearly and accurately. Help with learning, writing, coding, brainstorming, and everyday tasks. Admit uncertainty when necessary. Do not invent facts or sources."
        : "You are Malie AI, a careful and helpful AI writing assistant for students and professionals. Be clear, constructive, and honest. Preserve user intent. Never claim you verified sources or facts when you have not.\n\nTask-specific instructions: " +
          taskInstructions[mode];

      const userInput = isChat
        ? text.trim()
        : `Requested tone: ${tone}
Requested length: ${length}
Audience: ${audience}

User content and instructions:
${text}`;

      const model =
        env.GEMINI_MODEL || "gemini-2.5-flash";

      const apiUrl =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        encodeURIComponent(model) +
        ":generateContent";

      const payload = {
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
          maxOutputTokens: 3500
        }
      };

      let response;
      let lastNetworkError;

      // One initial attempt and at most one retry.
      for (let attempt = 1; attempt <= 2; attempt++) {
        const attemptStarted = Date.now();

        try {
          response = await fetch(apiUrl, {
            method: "POST",
            headers: {
              "content-type": "application/json",
              "x-goog-api-key": env.GEMINI_API_KEY
            },
            body: JSON.stringify(payload),
            signal: AbortSignal.timeout(25000)
          });

          console.log("Gemini request completed:", {
            status: response.status,
            attempt,
            durationMs: Date.now() - attemptStarted,
            model
          });

          if (
            transientStatuses.has(response.status) &&
            attempt < 2
          ) {
            await response.body?.cancel();

            console.warn(
              "Retrying temporary Gemini HTTP error:",
              {
                status: response.status,
                attempt
              }
            );

            await sleep(700);
            continue;
          }

          break;
        } catch (error) {
          lastNetworkError = error;

          console.error(
            "Gemini network/timeout error:",
            {
              attempt,
              durationMs: Date.now() - attemptStarted,
              errorName: error?.name || "Unknown",
              message: error?.message || "Unknown error"
            }
          );

          if (attempt < 2) {
            await sleep(700);
            continue;
          }
        }
      }

      if (!response) {
        console.error(
          "Gemini request failed after retries:",
          {
            durationMs: Date.now() - startedAt,
            errorName:
              lastNetworkError?.name || "Unknown"
          }
        );

        return respond({
          error:
            "Gemini could not be reached after retrying. Please try again."
        }, 502);
      }

      let data;
      const raw = await response.text();

      try {
        data = raw ? JSON.parse(raw) : {};
      } catch {
        console.error(
          "Gemini returned non-JSON data:",
          {
            status: response.status,
            durationMs: Date.now() - startedAt
          }
        );

        return respond({
          error:
            "Gemini returned an unexpected response. Please try again."
        }, 502);
      }

      if (!response.ok) {
        console.error(
          "Gemini API error after retry handling:",
          {
            httpStatus: response.status,
            statusName:
              data?.error?.status || "UNKNOWN",
            message:
              data?.error?.message ||
              "No error message returned",
            durationMs: Date.now() - startedAt,
            model
          }
        );

        if (response.status === 429) {
          return respond({
            error:
              "Gemini rate limit or quota reached. Check your Google AI usage and try again later."
          }, 429);
        }

        if (
          response.status === 400 ||
          response.status === 401 ||
          response.status === 403 ||
          response.status === 404
        ) {
          return respond({
            error:
              "Gemini rejected the request (HTTP " +
              response.status +
              "). Check the Worker logs for the exact reason."
          }, 502);
        }

        if (transientStatuses.has(response.status)) {
          return respond({
            error:
              "Gemini is temporarily unavailable (HTTP " +
              response.status +
              "). Please try again shortly."
          }, 503);
        }

        return respond({
          error:
            "Gemini returned an unexpected error (HTTP " +
            response.status +
            ")."
        }, 502);
      }

      const parts =
        data?.candidates?.[0]?.content?.parts;

      const result = Array.isArray(parts)
        ? parts
            .map(part =>
              typeof part.text === "string"
                ? part.text
                : ""
            )
            .filter(Boolean)
            .join("")
            .trim()
        : "";

      if (!result) {
        console.error(
          "Gemini response contained no text:",
          {
            finishReason:
              data?.candidates?.[0]?.finishReason ||
              "UNKNOWN",
            promptBlocked: Boolean(
              data?.promptFeedback?.blockReason
            ),
            durationMs: Date.now() - startedAt
          }
        );

        return respond({
          error:
            "Gemini returned no text. Please try a different prompt."
        }, 502);
      }

      console.log("Malie AI response completed:", {
        endpoint: path,
        durationMs: Date.now() - startedAt,
        outputCharacters: result.length
      });

      return isChat
        ? respond({ reply: result })
        : respond({ result });
    } catch (error) {
      console.error("Unhandled Worker exception:", {
        name: error?.name || "Unknown",
        message: error?.message || "Unknown error"
      });

      return respond({
        error:
          "Malie AI encountered an unexpected server error. Please try again."
      }, 500);
    }
  }
};
