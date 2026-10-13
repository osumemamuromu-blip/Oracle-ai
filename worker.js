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

  if (typeof text === "string") return text.trim();

  return "";
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
      throw new Error("The AI did not return valid quiz JSON.");
    }

    parsed = JSON.parse(cleaned.slice(start, end + 1));
  }

  const questions = Array.isArray(parsed)
    ? parsed
    : parsed.questions;

  if (!Array.isArray(questions) || !questions.length) {
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
      String(
        typeof option === "object"
          ? option.text ?? option.label ?? ""
          : option
      ).trim()
    );

    let answer = item.answerIndex;

    if (answer === undefined) answer = item.correctIndex;
    if (answer === undefined) answer = item.correctAnswerIndex;
    if (answer === undefined) {
      answer = item.correctAnswer ?? item.answer;
    }

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
        `Quiz question ${index + 1} has an invalid answer.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex: answer,
      explanation: String(item.explanation || "")
    };
  });

  return { questions: validated };
}

async function callAI(env, prompt, jsonMode = false) {
  if (!env.AI || typeof env.AI.run !== "function") {
    const error = new Error(
      "Cloudflare Workers AI is not connected. Check the AI binding."
    );
    error.status = 500;
    throw error;
  }

  const model = "@cf/google/gemma-4-26b-a4b-it";

  let result;

  try {
    result = await env.AI.run(model, {
      messages: [
        {
          role: "system",
          content: jsonMode
            ? "You are Malie AI, an accurate educational assistant. Follow the user's instructions and return only valid JSON, without Markdown fences or extra text."
            : "You are Malie AI, a helpful and accurate assistant. Follow the user's instructions, preserve supplied facts, and do not invent citations."
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
      /limit|quota|neuron/i.test(message)
        ? "Cloudflare AI's free usage limit may have been reached. Please try again after the daily reset."
        : /capacity|busy|overload/i.test(message)
        ? "Cloudflare AI is busy right now. Please try again shortly."
        : "Cloudflare AI could not complete the request. Please try again."
    );

    error.status = /limit|quota|neuron/i.test(message) ? 429 : 503;
    throw error;
  }

  const text = extractText(result);

  if (!text) {
    console.error("Unexpected Workers AI response:", JSON.stringify(result).slice(0, 1000));
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
        return respond({ error: "Please send valid JSON." }, 400);
      }

      try {
        if (path === "/api/chat") {
          const message = String(body.message || "").trim();

          if (!message) {
            return respond({ error: "Please enter a message." }, 400);
          }

          if (message.length > 30000) {
            return respond({ error: "Your message is too long." }, 400);
          }

          const reply = await callAI(
            env,
            "Answer the user's message clearly and honestly.\n\nUser message:\n" + message
          );

          return respond({ reply });
        }

        if (path === "/api/write") {
          const mode = String(body.mode || "correct").trim();
          const inputText = String(body.text || "").trim();
          const options = body.options || {};

          if (!inputText) {
            return respond({ error: "Please enter some text." }, 400);
          }

          if (inputText.length > 30000) {
            return respond({
              error: "Your text is too long. Please shorten it and try again."
            }, 400);
          }

          const instruction =
            taskInstructions[mode] || taskInstructions.correct;

          const result = await callAI(
            env,
            "Task: " + instruction +
            "\nUser options: " + JSON.stringify(options) +
            "\n\nText to process:\n" + inputText
          );

          return respond({ result });
        }

        if (path === "/api/quiz") {
          const inputText = String(body.text || "").trim();

          const count = Math.max(
            1,
            Math.min(20, Number.parseInt(body.count, 10) || 5)
          );

          const difficulty = String(body.difficulty || "medium").trim();

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

          const prompt =
            "Create exactly " + count +
            " multiple-choice questions based ONLY on the source material below.\n" +
            "Difficulty: " + difficulty + ".\n" +
            "Each question must have four answer options, one correct answer, " +
            "and a short explanation. Avoid duplicates and unsupported facts.\n" +
            "Return ONLY valid JSON in this structure:\n" +
            '{"questions":[{"question":"Question text","options":["A","B","C","D"],"answerIndex":0,"explanation":"Short explanation"}]}' +
            "\n\nSource material:\n" + inputText;

          const resultText = await callAI(env, prompt, true);
          const quiz = parseQuiz(resultText);

          if (quiz.questions.length !== count) {
            throw new Error(
              `The AI returned ${quiz.questions.length} questions instead of ${count}. Please try again.`
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
      return respond({ error: "API endpoint not found." }, 404);
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Malie AI is running.", {
      headers: corsHeaders
    });
  }
};
