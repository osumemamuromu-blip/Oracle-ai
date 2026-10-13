
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
    "Rewrite the supplied text in the requested tone for the intended audience. Preserve its meaning and key details. Do not add unsupported claims. Return the rewritten version only unless a brief note is needed.",

  rewrite:
    "Rewrite the supplied text to improve its expression and readability. Preserve the original meaning, important facts, and intent. Do not add unsupported claims.",

  expand:
    "Expand the supplied text with relevant explanations and detail while preserving its original meaning. Do not invent factual claims, examples presented as facts, or citations."
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

function extractText(data) {
  const parts = data?.candidates?.[0]?.content?.parts;

  return Array.isArray(parts)
    ? parts
        .map(part =>
          typeof part.text === "string" ? part.text : ""
        )
        .filter(Boolean)
        .join("")
        .trim()
    : "";
}

function parseQuiz(text) {
  let source = text.trim();

  // Remove JSON code fences if the model includes them.
  source = source
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();

  let parsed;

  try {
    parsed = JSON.parse(source);
  } catch {
    // Try extracting a JSON object from surrounding text.
    const start = source.indexOf("{");
    const end = source.lastIndexOf("}");

    if (start === -1 || end <= start) {
      throw new Error("The AI did not return valid quiz JSON.");
    }

    parsed = JSON.parse(source.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed?.questions;

  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error("The AI returned no quiz questions.");
  }

  const validQuestions = questions.map((item, index) => {
    const question = String(
      item?.question || item?.prompt || ""
    ).trim();

    const options = item?.options || item?.choices;

    if (
      !question ||
      !Array.isArray(options) ||
      options.length < 2
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid format.`
      );
    }

    const cleanOptions = options.map(option =>
      String(option).trim()
    );

    let answerIndex = item?.answerIndex;

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctAnswerIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      const answer = item?.correctAnswer ?? item?.answer;

      if (typeof answer === "number") {
        answerIndex = answer;
      } else if (typeof answer === "string") {
        const letter = answer.trim().match(/^([A-D])(?:[.)\s]|$)/i);

        if (letter) {
          answerIndex = letter[1].toUpperCase().charCodeAt(0) - 65;
        } else {
          answerIndex = cleanOptions.findIndex(
            option => option.toLowerCase() === answer.trim().toLowerCase()
          );
        }
      }
    }

    if (
      !Number.isInteger(answerIndex) ||
      answerIndex < 0 ||
      answerIndex >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has no valid correct answer.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex,
      explanation: String(item?.explanation || "")
    };
  });

  return validQuestions;
}

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
          endpoints: [
            "/api/chat",
            "/api/write",
            "/api/quiz"
          ]
        });
      }

      const isChat = path === "/api/chat";
      const isWriting = path === "/api/write";
      const isQuiz = path === "/api/quiz";

      if (!isChat && !isWriting && !isQuiz) {
        if (
          env.ASSETS &&
          typeof env.ASSETS.fetch === "function"
        ) {
          return env.ASSETS.fetch(request);
        }

        return respond({ error: "Route not found." }, 404);
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
        return respond({ error: "Invalid JSON request." }, 400);
      }

      const mode = isChat ? "generate" : body?.mode;
      const text = isChat ? body?.message : body?.text;
      const options = body?.options || {};

      if (isWriting &&
          !Object.prototype.hasOwnProperty.call(taskInstructions, mode)) {
        return respond({ error: "Choose a valid writing mode." }, 400);
      }

      if (typeof text !== "string" || !text.trim()) {
        return respond({
          error: isChat
            ? "Please enter a message."
            : isQuiz
              ? "Enter some study text before generating a quiz."
              : "Enter text or instructions first."
        }, 400);
      }

      if (text.length > 30000) {
        return respond({
          error: "Please keep each request under 30,000 characters."
        }, 413);
      }

      let systemPrompt;
      let userInput;
      let generationConfig = {
        maxOutputTokens: 3500
      };

      if (isQuiz) {
        const requestedCount = Number(body?.count ?? 5);
        const count = Number.isFinite(requestedCount)
          ? Math.max(1, Math.min(20, Math.floor(requestedCount)))
          : 5;

        const allowedDifficulties = new Set([
          "easy",
          "medium",
          "hard",
          "mixed"
        ]);

        const difficulty = String(
          body?.difficulty || "medium"
        ).toLowerCase();

        const safeDifficulty = allowedDifficulties.has(difficulty)
          ? difficulty
          : "medium";

        systemPrompt =
          "You are Malie AI, a careful educational quiz generator. " +
          "Use only the supplied study material to create questions. " +
          "Do not invent claims that are not supported by the material. " +
          "Return valid JSON only, with no Markdown fences or surrounding commentary. " +
          'Use this exact structure: {"questions":[{"question":"Question text","options":["Option A","Option B","Option C","Option D"],"answerIndex":0,"explanation":"Brief explanation based on the material"}]}. ' +
          "answerIndex must be a zero-based integer matching the correct option. " +
          "Each question must have four distinct options and exactly one correct answer. " +
          "Make distractors plausible but clearly incorrect according to the material. " +
          "Do not repeat questions.";

        userInput =
          `Generate exactly ${count} multiple-choice questions.\n` +
          `Difficulty: ${safeDifficulty}.\n` +
          "Base every question and answer on the following study material.\n\n" +
          "STUDY MATERIAL:\n" +
          text.trim();

        generationConfig = {
          maxOutputTokens: 6000,
          responseMimeType: "application/json"
        };
      } else {
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

        systemPrompt = isChat
          ? "You are Malie AI, a helpful general-purpose AI assistant. Answer questions clearly and accurately. Help with learning, writing, coding, brainstorming, and everyday tasks. Admit uncertainty when necessary. Do not invent facts or sources."
          : "You are Malie AI, a careful and helpful AI writing assistant for students and professionals. Be clear, constructive, and honest. Preserve user intent. Never claim you verified sources or facts when you have not.\n\nTask-specific instructions: " +
            taskInstructions[mode];

        userInput = isChat
          ? text.trim()
          : `Requested tone: ${tone}
Requested length: ${length}
Audience: ${audience}

User content and instructions:
${text}`;
      }

      const model = env.GEMINI_MODEL || "gemini-2.5-flash";

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
        generationConfig
      };

      let response;
      let lastNetworkError;

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
            model,
            endpoint: path
          });

          if (
            transientStatuses.has(response.status) &&
            attempt < 2
          ) {
            await response.body?.cancel();
            await sleep(700);
            continue;
          }

          break;
        } catch (error) {
          lastNetworkError = error;

          console.error("Gemini network/timeout error:", {
            attempt,
            durationMs: Date.now() - attemptStarted,
            errorName: error?.name || "Unknown"
          });

          if (attempt < 2) {
            await sleep(700);
          }
        }
      }

      if (!response) {
        console.error("Gemini request failed after retries:", {
          durationMs: Date.now() - startedAt,
          errorName: lastNetworkError?.name || "Unknown"
        });

        return respond({
          error: "Gemini could not be reached after retrying. Please try again."
        }, 502);
      }

      const raw = await response.text();
      let data;

      try {
        data = raw ? JSON.parse(raw) : {};
      } catch {
        console.error("Gemini returned non-JSON data:", {
          status: response.status,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned an unexpected response. Please try again."
        }, 502);
      }

      if (!response.ok) {
        console.error("Gemini API error:", {
          httpStatus: response.status,
          statusName: data?.error?.status || "UNKNOWN",
          message: data?.error?.message || "No error message returned",
          durationMs: Date.now() - startedAt,
          model,
          endpoint: path
        });

        if (response.status === 429) {
          return respond({
            error: "Gemini rate limit or quota reached. Check your Google AI usage and try again later."
          }, 429);
        }

        if ([400, 401, 403, 404].includes(response.status)) {
          return respond({
            error: "Gemini rejected the request (HTTP " +
              response.status +
              "). Check the Worker logs for the exact reason."
          }, 502);
        }

        if (transientStatuses.has(response.status)) {
          return respond({
            error: "Gemini is temporarily unavailable (HTTP " +
              response.status +
              "). Please try again shortly."
          }, 503);
        }

        return respond({
          error: "Gemini returned an unexpected error (HTTP " +
            response.status +
            ")."
        }, 502);
      }

      const result = extractText(data);

      if (!result) {
        console.error("Gemini response contained no text:", {
          finishReason: data?.candidates?.[0]?.finishReason || "UNKNOWN",
          promptBlocked: Boolean(data?.promptFeedback?.blockReason),
          endpoint: path,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned no text. Please try a different prompt."
        }, 502);
      }

      if (isQuiz) {
        try {
          const questions = parseQuiz(result);

          console.log("Quiz generated:", {
            questionCount: questions.length,
            durationMs: Date.now() - startedAt
          });

          return respond({ questions });
        } catch (error) {
          console.error("Quiz parsing failed:", {
            message: error?.message || "Unknown parsing error",
            durationMs: Date.now() - startedAt
          });

          return respond({
            error: "The AI response could not be formatted as a quiz. Please try again."
          }, 502);
        }
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
        error: "Malie AI encountered an unexpected server error. Please try again."
      }, 500);
    }
  }
};
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
    "Rewrite the supplied text in the requested tone for the intended audience. Preserve its meaning and key details. Do not add unsupported claims. Return the rewritten version only unless a brief note is needed.",

  rewrite:
    "Rewrite the supplied text to improve its expression and readability. Preserve the original meaning, important facts, and intent. Do not add unsupported claims.",

  expand:
    "Expand the supplied text with relevant explanations and detail while preserving its original meaning. Do not invent factual claims, examples presented as facts, or citations."
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

function extractText(data) {
  const parts = data?.candidates?.[0]?.content?.parts;

  return Array.isArray(parts)
    ? parts
        .map(part =>
          typeof part.text === "string" ? part.text : ""
        )
        .filter(Boolean)
        .join("")
        .trim()
    : "";
}

function parseQuiz(text) {
  let source = text.trim();

  // Remove JSON code fences if the model includes them.
  source = source
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();

  let parsed;

  try {
    parsed = JSON.parse(source);
  } catch {
    // Try extracting a JSON object from surrounding text.
    const start = source.indexOf("{");
    const end = source.lastIndexOf("}");

    if (start === -1 || end <= start) {
      throw new Error("The AI did not return valid quiz JSON.");
    }

    parsed = JSON.parse(source.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed?.questions;

  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error("The AI returned no quiz questions.");
  }

  const validQuestions = questions.map((item, index) => {
    const question = String(
      item?.question || item?.prompt || ""
    ).trim();

    const options = item?.options || item?.choices;

    if (
      !question ||
      !Array.isArray(options) ||
      options.length < 2
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid format.`
      );
    }

    const cleanOptions = options.map(option =>
      String(option).trim()
    );

    let answerIndex = item?.answerIndex;

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctAnswerIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      const answer = item?.correctAnswer ?? item?.answer;

      if (typeof answer === "number") {
        answerIndex = answer;
      } else if (typeof answer === "string") {
        const letter = answer.trim().match(/^([A-D])(?:[.)\s]|$)/i);

        if (letter) {
          answerIndex = letter[1].toUpperCase().charCodeAt(0) - 65;
        } else {
          answerIndex = cleanOptions.findIndex(
            option => option.toLowerCase() === answer.trim().toLowerCase()
          );
        }
      }
    }

    if (
      !Number.isInteger(answerIndex) ||
      answerIndex < 0 ||
      answerIndex >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has no valid correct answer.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex,
      explanation: String(item?.explanation || "")
    };
  });

  return validQuestions;
}

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
          endpoints: [
            "/api/chat",
            "/api/write",
            "/api/quiz"
          ]
        });
      }

      const isChat = path === "/api/chat";
      const isWriting = path === "/api/write";
      const isQuiz = path === "/api/quiz";

      if (!isChat && !isWriting && !isQuiz) {
        if (
          env.ASSETS &&
          typeof env.ASSETS.fetch === "function"
        ) {
          return env.ASSETS.fetch(request);
        }

        return respond({ error: "Route not found." }, 404);
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
        return respond({ error: "Invalid JSON request." }, 400);
      }

      const mode = isChat ? "generate" : body?.mode;
      const text = isChat ? body?.message : body?.text;
      const options = body?.options || {};

      if (isWriting &&
          !Object.prototype.hasOwnProperty.call(taskInstructions, mode)) {
        return respond({ error: "Choose a valid writing mode." }, 400);
      }

      if (typeof text !== "string" || !text.trim()) {
        return respond({
          error: isChat
            ? "Please enter a message."
            : isQuiz
              ? "Enter some study text before generating a quiz."
              : "Enter text or instructions first."
        }, 400);
      }

      if (text.length > 30000) {
        return respond({
          error: "Please keep each request under 30,000 characters."
        }, 413);
      }

      let systemPrompt;
      let userInput;
      let generationConfig = {
        maxOutputTokens: 3500
      };

      if (isQuiz) {
        const requestedCount = Number(body?.count ?? 5);
        const count = Number.isFinite(requestedCount)
          ? Math.max(1, Math.min(20, Math.floor(requestedCount)))
          : 5;

        const allowedDifficulties = new Set([
          "easy",
          "medium",
          "hard",
          "mixed"
        ]);

        const difficulty = String(
          body?.difficulty || "medium"
        ).toLowerCase();

        const safeDifficulty = allowedDifficulties.has(difficulty)
          ? difficulty
          : "medium";

        systemPrompt =
          "You are Malie AI, a careful educational quiz generator. " +
          "Use only the supplied study material to create questions. " +
          "Do not invent claims that are not supported by the material. " +
          "Return valid JSON only, with no Markdown fences or surrounding commentary. " +
          'Use this exact structure: {"questions":[{"question":"Question text","options":["Option A","Option B","Option C","Option D"],"answerIndex":0,"explanation":"Brief explanation based on the material"}]}. ' +
          "answerIndex must be a zero-based integer matching the correct option. " +
          "Each question must have four distinct options and exactly one correct answer. " +
          "Make distractors plausible but clearly incorrect according to the material. " +
          "Do not repeat questions.";

        userInput =
          `Generate exactly ${count} multiple-choice questions.\n` +
          `Difficulty: ${safeDifficulty}.\n` +
          "Base every question and answer on the following study material.\n\n" +
          "STUDY MATERIAL:\n" +
          text.trim();

        generationConfig = {
          maxOutputTokens: 6000,
          responseMimeType: "application/json"
        };
      } else {
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

        systemPrompt = isChat
          ? "You are Malie AI, a helpful general-purpose AI assistant. Answer questions clearly and accurately. Help with learning, writing, coding, brainstorming, and everyday tasks. Admit uncertainty when necessary. Do not invent facts or sources."
          : "You are Malie AI, a careful and helpful AI writing assistant for students and professionals. Be clear, constructive, and honest. Preserve user intent. Never claim you verified sources or facts when you have not.\n\nTask-specific instructions: " +
            taskInstructions[mode];

        userInput = isChat
          ? text.trim()
          : `Requested tone: ${tone}
Requested length: ${length}
Audience: ${audience}

User content and instructions:
${text}`;
      }

      const model = env.GEMINI_MODEL || "gemini-2.5-flash";

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
        generationConfig
      };

      let response;
      let lastNetworkError;

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
            model,
            endpoint: path
          });

          if (
            transientStatuses.has(response.status) &&
            attempt < 2
          ) {
            await response.body?.cancel();
            await sleep(700);
            continue;
          }

          break;
        } catch (error) {
          lastNetworkError = error;

          console.error("Gemini network/timeout error:", {
            attempt,
            durationMs: Date.now() - attemptStarted,
            errorName: error?.name || "Unknown"
          });

          if (attempt < 2) {
            await sleep(700);
          }
        }
      }

      if (!response) {
        console.error("Gemini request failed after retries:", {
          durationMs: Date.now() - startedAt,
          errorName: lastNetworkError?.name || "Unknown"
        });

        return respond({
          error: "Gemini could not be reached after retrying. Please try again."
        }, 502);
      }

      const raw = await response.text();
      let data;

      try {
        data = raw ? JSON.parse(raw) : {};
      } catch {
        console.error("Gemini returned non-JSON data:", {
          status: response.status,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned an unexpected response. Please try again."
        }, 502);
      }

      if (!response.ok) {
        console.error("Gemini API error:", {
          httpStatus: response.status,
          statusName: data?.error?.status || "UNKNOWN",
          message: data?.error?.message || "No error message returned",
          durationMs: Date.now() - startedAt,
          model,
          endpoint: path
        });

        if (response.status === 429) {
          return respond({
            error: "Gemini rate limit or quota reached. Check your Google AI usage and try again later."
          }, 429);
        }

        if ([400, 401, 403, 404].includes(response.status)) {
          return respond({
            error: "Gemini rejected the request (HTTP " +
              response.status +
              "). Check the Worker logs for the exact reason."
          }, 502);
        }

        if (transientStatuses.has(response.status)) {
          return respond({
            error: "Gemini is temporarily unavailable (HTTP " +
              response.status +
              "). Please try again shortly."
          }, 503);
        }

        return respond({
          error: "Gemini returned an unexpected error (HTTP " +
            response.status +
            ")."
        }, 502);
      }

      const result = extractText(data);

      if (!result) {
        console.error("Gemini response contained no text:", {
          finishReason: data?.candidates?.[0]?.finishReason || "UNKNOWN",
          promptBlocked: Boolean(data?.promptFeedback?.blockReason),
          endpoint: path,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned no text. Please try a different prompt."
        }, 502);
      }

      if (isQuiz) {
        try {
          const questions = parseQuiz(result);

          console.log("Quiz generated:", {
            questionCount: questions.length,
            durationMs: Date.now() - startedAt
          });

          return respond({ questions });
        } catch (error) {
          console.error("Quiz parsing failed:", {
            message: error?.message || "Unknown parsing error",
            durationMs: Date.now() - startedAt
          });

          return respond({
            error: "The AI response could not be formatted as a quiz. Please try again."
          }, 502);
        }
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
        error: "Malie AI encountered an unexpected server error. Please try again."
      }, 500);
    }
  }
};
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
    "Rewrite the supplied text in the requested tone for the intended audience. Preserve its meaning and key details. Do not add unsupported claims. Return the rewritten version only unless a brief note is needed.",

  rewrite:
    "Rewrite the supplied text to improve its expression and readability. Preserve the original meaning, important facts, and intent. Do not add unsupported claims.",

  expand:
    "Expand the supplied text with relevant explanations and detail while preserving its original meaning. Do not invent factual claims, examples presented as facts, or citations."
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

function extractText(data) {
  const parts = data?.candidates?.[0]?.content?.parts;

  return Array.isArray(parts)
    ? parts
        .map(part =>
          typeof part.text === "string" ? part.text : ""
        )
        .filter(Boolean)
        .join("")
        .trim()
    : "";
}

function parseQuiz(text) {
  let source = text.trim();

  // Remove JSON code fences if the model includes them.
  source = source
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();

  let parsed;

  try {
    parsed = JSON.parse(source);
  } catch {
    // Try extracting a JSON object from surrounding text.
    const start = source.indexOf("{");
    const end = source.lastIndexOf("}");

    if (start === -1 || end <= start) {
      throw new Error("The AI did not return valid quiz JSON.");
    }

    parsed = JSON.parse(source.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed?.questions;

  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error("The AI returned no quiz questions.");
  }

  const validQuestions = questions.map((item, index) => {
    const question = String(
      item?.question || item?.prompt || ""
    ).trim();

    const options = item?.options || item?.choices;

    if (
      !question ||
      !Array.isArray(options) ||
      options.length < 2
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid format.`
      );
    }

    const cleanOptions = options.map(option =>
      String(option).trim()
    );

    let answerIndex = item?.answerIndex;

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      answerIndex = item?.correctAnswerIndex;
    }

    if (!Number.isInteger(answerIndex)) {
      const answer = item?.correctAnswer ?? item?.answer;

      if (typeof answer === "number") {
        answerIndex = answer;
      } else if (typeof answer === "string") {
        const letter = answer.trim().match(/^([A-D])(?:[.)\s]|$)/i);

        if (letter) {
          answerIndex = letter[1].toUpperCase().charCodeAt(0) - 65;
        } else {
          answerIndex = cleanOptions.findIndex(
            option => option.toLowerCase() === answer.trim().toLowerCase()
          );
        }
      }
    }

    if (
      !Number.isInteger(answerIndex) ||
      answerIndex < 0 ||
      answerIndex >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has no valid correct answer.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex,
      explanation: String(item?.explanation || "")
    };
  });

  return validQuestions;
}

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
          endpoints: [
            "/api/chat",
            "/api/write",
            "/api/quiz"
          ]
        });
      }

      const isChat = path === "/api/chat";
      const isWriting = path === "/api/write";
      const isQuiz = path === "/api/quiz";

      if (!isChat && !isWriting && !isQuiz) {
        if (
          env.ASSETS &&
          typeof env.ASSETS.fetch === "function"
        ) {
          return env.ASSETS.fetch(request);
        }

        return respond({ error: "Route not found." }, 404);
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
        return respond({ error: "Invalid JSON request." }, 400);
      }

      const mode = isChat ? "generate" : body?.mode;
      const text = isChat ? body?.message : body?.text;
      const options = body?.options || {};

      if (isWriting &&
          !Object.prototype.hasOwnProperty.call(taskInstructions, mode)) {
        return respond({ error: "Choose a valid writing mode." }, 400);
      }

      if (typeof text !== "string" || !text.trim()) {
        return respond({
          error: isChat
            ? "Please enter a message."
            : isQuiz
              ? "Enter some study text before generating a quiz."
              : "Enter text or instructions first."
        }, 400);
      }

      if (text.length > 30000) {
        return respond({
          error: "Please keep each request under 30,000 characters."
        }, 413);
      }

      let systemPrompt;
      let userInput;
      let generationConfig = {
        maxOutputTokens: 3500
      };

      if (isQuiz) {
        const requestedCount = Number(body?.count ?? 5);
        const count = Number.isFinite(requestedCount)
          ? Math.max(1, Math.min(20, Math.floor(requestedCount)))
          : 5;

        const allowedDifficulties = new Set([
          "easy",
          "medium",
          "hard",
          "mixed"
        ]);

        const difficulty = String(
          body?.difficulty || "medium"
        ).toLowerCase();

        const safeDifficulty = allowedDifficulties.has(difficulty)
          ? difficulty
          : "medium";

        systemPrompt =
          "You are Malie AI, a careful educational quiz generator. " +
          "Use only the supplied study material to create questions. " +
          "Do not invent claims that are not supported by the material. " +
          "Return valid JSON only, with no Markdown fences or surrounding commentary. " +
          'Use this exact structure: {"questions":[{"question":"Question text","options":["Option A","Option B","Option C","Option D"],"answerIndex":0,"explanation":"Brief explanation based on the material"}]}. ' +
          "answerIndex must be a zero-based integer matching the correct option. " +
          "Each question must have four distinct options and exactly one correct answer. " +
          "Make distractors plausible but clearly incorrect according to the material. " +
          "Do not repeat questions.";

        userInput =
          `Generate exactly ${count} multiple-choice questions.\n` +
          `Difficulty: ${safeDifficulty}.\n` +
          "Base every question and answer on the following study material.\n\n" +
          "STUDY MATERIAL:\n" +
          text.trim();

        generationConfig = {
          maxOutputTokens: 6000,
          responseMimeType: "application/json"
        };
      } else {
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

        systemPrompt = isChat
          ? "You are Malie AI, a helpful general-purpose AI assistant. Answer questions clearly and accurately. Help with learning, writing, coding, brainstorming, and everyday tasks. Admit uncertainty when necessary. Do not invent facts or sources."
          : "You are Malie AI, a careful and helpful AI writing assistant for students and professionals. Be clear, constructive, and honest. Preserve user intent. Never claim you verified sources or facts when you have not.\n\nTask-specific instructions: " +
            taskInstructions[mode];

        userInput = isChat
          ? text.trim()
          : `Requested tone: ${tone}
Requested length: ${length}
Audience: ${audience}

User content and instructions:
${text}`;
      }

      const model = env.GEMINI_MODEL || "gemini-2.5-flash";

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
        generationConfig
      };

      let response;
      let lastNetworkError;

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
            model,
            endpoint: path
          });

          if (
            transientStatuses.has(response.status) &&
            attempt < 2
          ) {
            await response.body?.cancel();
            await sleep(700);
            continue;
          }

          break;
        } catch (error) {
          lastNetworkError = error;

          console.error("Gemini network/timeout error:", {
            attempt,
            durationMs: Date.now() - attemptStarted,
            errorName: error?.name || "Unknown"
          });

          if (attempt < 2) {
            await sleep(700);
          }
        }
      }

      if (!response) {
        console.error("Gemini request failed after retries:", {
          durationMs: Date.now() - startedAt,
          errorName: lastNetworkError?.name || "Unknown"
        });

        return respond({
          error: "Gemini could not be reached after retrying. Please try again."
        }, 502);
      }

      const raw = await response.text();
      let data;

      try {
        data = raw ? JSON.parse(raw) : {};
      } catch {
        console.error("Gemini returned non-JSON data:", {
          status: response.status,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned an unexpected response. Please try again."
        }, 502);
      }

      if (!response.ok) {
        console.error("Gemini API error:", {
          httpStatus: response.status,
          statusName: data?.error?.status || "UNKNOWN",
          message: data?.error?.message || "No error message returned",
          durationMs: Date.now() - startedAt,
          model,
          endpoint: path
        });

        if (response.status === 429) {
          return respond({
            error: "Gemini rate limit or quota reached. Check your Google AI usage and try again later."
          }, 429);
        }

        if ([400, 401, 403, 404].includes(response.status)) {
          return respond({
            error: "Gemini rejected the request (HTTP " +
              response.status +
              "). Check the Worker logs for the exact reason."
          }, 502);
        }

        if (transientStatuses.has(response.status)) {
          return respond({
            error: "Gemini is temporarily unavailable (HTTP " +
              response.status +
              "). Please try again shortly."
          }, 503);
        }

        return respond({
          error: "Gemini returned an unexpected error (HTTP " +
            response.status +
            ")."
        }, 502);
      }

      const result = extractText(data);

      if (!result) {
        console.error("Gemini response contained no text:", {
          finishReason: data?.candidates?.[0]?.finishReason || "UNKNOWN",
          promptBlocked: Boolean(data?.promptFeedback?.blockReason),
          endpoint: path,
          durationMs: Date.now() - startedAt
        });

        return respond({
          error: "Gemini returned no text. Please try a different prompt."
        }, 502);
      }

      if (isQuiz) {
        try {
          const questions = parseQuiz(result);

          console.log("Quiz generated:", {
            questionCount: questions.length,
            durationMs: Date.now() - startedAt
          });

          return respond({ questions });
        } catch (error) {
          console.error("Quiz parsing failed:", {
            message: error?.message || "Unknown parsing error",
            durationMs: Date.now() - startedAt
          });

          return respond({
            error: "The AI response could not be formatted as a quiz. Please try again."
          }, 502);
        }
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
        error: "Malie AI encountered an unexpected server error. Please try again."
      }, 500);
    }
  }
};
