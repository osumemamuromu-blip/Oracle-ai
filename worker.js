const taskInstructions = {
  correct:
    "Correct grammar, spelling, punctuation, and sentence errors. Preserve the writer's meaning, facts, and voice. Return only the corrected text unless a brief note is necessary.",

  improve:
    "Improve clarity, grammar, flow, and readability while preserving the writer's meaning, facts, and voice. Return the improved text only.",

  generate:
    "Write original, clear, useful content based on the user's instructions. Follow the requested format and tone.",

  explain:
    "Explain the provided material clearly and step by step. Use simple language and examples where helpful.",

  summarize:
    "Summarize the supplied material accurately. Preserve its main ideas and important details. Do not invent information.",

  tone:
    "Rewrite the supplied text in the requested tone while preserving its meaning and important details.",

  rewrite:
    "Rewrite the supplied text to improve clarity and natural flow while preserving its meaning and facts.",

  expand:
    "Expand the supplied text with relevant explanations and details. Do not invent facts or citations."
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization"
};

const transientStatuses = [500, 502, 503, 504];

function respond(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json; charset=utf-8"
    }
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function extractText(data) {
  return (data?.candidates?.[0]?.content?.parts || [])
    .map(part => part.text || "")
    .join("")
    .trim();
}

function parseQuiz(text) {
  let cleaned = text.trim();

  cleaned = cleaned
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "");

  let parsed;

  try {
    parsed = JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");

    if (start === -1 || end <= start) {
      throw new Error("The AI did not return valid quiz JSON.");
    }

    parsed = JSON.parse(cleaned.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed.questions;

  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error("No quiz questions were returned.");
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
        `Quiz question ${index + 1} is incomplete.`
      );
    }

    const cleanOptions = options.map(option =>
      String(option).trim()
    );

    let answerIndex = item.answerIndex;

    if (answerIndex === undefined) {
      answerIndex = item.correctIndex;
    }

    if (answerIndex === undefined) {
      answerIndex = item.correctAnswerIndex;
    }

    if (answerIndex === undefined) {
      answerIndex = item.correctAnswer ?? item.answer;
    }

    if (typeof answerIndex === "string") {
      const answerText = answerIndex.trim();

      const letterIndex = answerText
        .toUpperCase()
        .match(/^[A-D]$/);

      if (letterIndex) {
        answerIndex =
          letterIndex[0].charCodeAt(0) - 65;
      } else {
        const found = cleanOptions.findIndex(
          option =>
            option.toLowerCase() ===
            answerText.toLowerCase()
        );

        answerIndex = found;
      }
    }

    answerIndex = Number(answerIndex);

    if (
      !Number.isInteger(answerIndex) ||
      answerIndex < 0 ||
      answerIndex >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid answer.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex,
      explanation: String(item.explanation || "")
    };
  });

  return { questions: validated };
}

async function callGemini(env, prompt, jsonMode = false) {
  if (!env.GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is not configured.");
  }

  const model = env.GEMINI_MODEL || "gemini-3.5-flash";

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent";

  const body = {
    contents: [
      {
        role: "user",
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: jsonMode ? 0.3 : 0.7,
      ...(jsonMode
        ? { responseMimeType: "application/json" }
        : {})
    }
  };

  let response;
  let lastError;

  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": env.GEMINI_API_KEY
        },
        body: JSON.stringify(body)
      });

      if (
        transientStatuses.includes(response.status) &&
        attempt === 0
      ) {
        await sleep(1000);
        continue;
      }

      break;
    } catch (error) {
      lastError = error;

      if (attempt === 0) {
        await sleep(1000);
        continue;
      }

      throw new Error(
        "Could not connect to Gemini. Please try again."
      );
    }
  }

  if (!response) {
    throw lastError || new Error("No response from Gemini.");
  }

  const raw = await response.text();

  let data;

  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error(
      `Gemini returned an unexpected response (HTTP ${response.status}).`
    );
  }

  if (!response.ok) {
    const apiMessage =
      data?.error?.message || "Gemini request failed.";

    const error = new Error(apiMessage);
    error.status = response.status;
    throw error;
  }

  const text = extractText(data);

  if (!text) {
    throw new Error(
      "Gemini returned no text. Please try again."
    );
  }

  return text;
}

function getErrorStatus(error) {
  if (error.status === 429) return 429;
  if (error.status === 400) return 400;
  if (error.status === 401 || error.status === 403) {
    return error.status;
  }
  return 500;
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
        endpoints: [
          "/api/health",
          "/api/chat",
          "/api/write",
          "/api/quiz"
        ]
      });
    }

    if (
      request.method === "POST" &&
      ["/api/chat", "/api/write", "/api/quiz"].includes(path)
    ) {
      let body;

      try {
        body = await request.json();
      } catch {
        return respond(
          { error: "Please send valid JSON." },
          400
        );
      }

      try {
        if (path === "/api/chat") {
          const message = String(body.message || "").trim();

          if (!message) {
            return respond(
              { error: "Please enter a message." },
              400
            );
          }

          const prompt =
            "You are Malie AI, a helpful and accurate AI assistant. " +
            "Answer the user's message clearly and honestly. " +
            "Do not invent facts or citations.\n\n" +
            "User message:\n" + message;

          const reply = await callGemini(env, prompt);

          return respond({ reply });
        }

        if (path === "/api/write") {
          const mode = String(body.mode || "correct").trim();
          const inputText = String(body.text || "").trim();
          const options = body.options || {};

          if (!inputText) {
            return respond(
              { error: "Please enter some text." },
              400
            );
          }

          if (inputText.length > 30000) {
            return respond(
              {
                error:
                  "Your text is too long. Please shorten it and try again."
              },
              400
            );
          }

          const instruction =
            taskInstructions[mode] || taskInstructions.correct;

          const prompt =
            "You are Malie AI, an expert writing assistant.\n" +
            "Task: " + instruction + "\n" +
            "User options: " + JSON.stringify(options) + "\n\n" +
            "Text to process:\n" + inputText;

          const result = await callGemini(env, prompt);

          return respond({ result });
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
            return respond(
              {
                error:
                  "Please provide notes or text to generate a quiz."
              },
              400
            );
          }

          if (inputText.length > 30000) {
            return respond(
              {
                error:
                  "The source text is too long. Please shorten it and try again."
              },
              400
            );
          }

          const prompt =
            "You are Malie AI, an educational quiz generator.\n" +
            "Create exactly " + count +
            " multiple-choice questions based ONLY on the source material below.\n" +
            "Difficulty: " + difficulty + ".\n" +
            "Each question must have four answer options, one correct answer, " +
            "and a short explanation.\n" +
            "Avoid duplicate questions. Do not introduce unsupported facts.\n" +
            "Return ONLY valid JSON in this exact structure:\n" +
            '{"questions":[{"question":"Question text","options":["A","B","C","D"],"answerIndex":0,"explanation":"Short explanation"}]}\n\n' +
            "Source material:\n" + inputText;

          const resultText = await callGemini(
            env,
            prompt,
            true
          );

          const quiz = parseQuiz(resultText);

          return respond(quiz);
        }
      } catch (error) {
        const status = getErrorStatus(error);

        console.error("API error:", {
          path,
          status,
          message: error.message
        });

        return respond(
          {
            error:
              status === 429
                ? "Gemini rate limit or quota reached. Please wait and try again, or check your Gemini API quota."
                : error.message || "The request failed. Please try again."
          },
          status
        );
      }
    }

    if (path.startsWith("/api/")) {
      return respond(
        { error: "API endpoint not found." },
        404
      );
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Malie AI is running.", {
      headers: corsHeaders
    });
  }
};
