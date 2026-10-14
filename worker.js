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

  let result;

  try {
    result = await env.AI.run(model, {
      messages: [
        {
          role: "system",
          content: jsonMode
            ? [
"You are Malie AI, the AI assistant created and developed by Campus Creative, founded by Osume Prosper.",
"IDENTITY RULE: When asked who created you, who made you, who developed you, who owns Malie AI, or who your founder is, answer clearly that Malie AI was created and developed by Campus Creative, founded by Osume Prosper.",
"Never answer that Google, Gemini, Google AI, Cloudflare, or another technology provider created Malie AI. They may provide underlying technology, but Campus Creative is the company behind the Malie AI product.",
"If asked who provides your underlying AI technology, explain that Malie AI currently runs on Cloudflare Workers AI and uses the configured Gemma model.",
"Be truthful, consistent, and confident about your identity. Do not confuse the creator of Malie AI with the provider of its underlying AI model.", ].join(" ")
            : [
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
              ].join(" ")
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
            return respond({
              error: "Please enter a message."
            }, 400);
          }

          if (message.length > 30000) {
            return respond({
              error: "Your message is too long. Please shorten it."
            }, 400);
          }
        
const identityQuestion =
  /\b(?:who\s+(?:created|made|built|developed|owns|founded)\s+(?:you|malie(?:\s+ai)?|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?(?:the\s+)?(?:creator|founder|owner)\s+of\s+(?:malie(?:\s+ai)?|you|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?your\s+(?:creator|founder|owner)|who\s+is\s+behind\s+(?:malie(?:\s+ai)?|you)|(?:are\s+you|is\s+malie(?:\s+ai)?)\s+(?:created|made|built|developed)\s+by)\b/i;

if (identityQuestion.test(message)) {
  return respond({
    reply:
      "I am Malie AI, developed by Campus Creative, founded by Osume Prosper. My underlying AI model is Google's Gemma, provided through Cloudflare Workers AI. Google provides the model technology; Campus Creative created the Malie AI product."
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
              "",
              "User message:",
              message
            ].join("\n")
          );

          return respond({ reply });
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
