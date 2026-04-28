"use strict";

// Set AI_PROVIDER in .env to either "google" or "openai".
// Both API keys can be present — this variable controls which one is used.

const provider = (process.env.AI_PROVIDER || "").toLowerCase();

let _generate = null;

if (provider === "google") {
  const { GoogleGenerativeAI } = require("@google/generative-ai");
  const genAI = new GoogleGenerativeAI(process.env.GOOGLE_API_KEY);
  const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });
  _generate = async (prompt) => {
    const result = await model.generateContent(prompt);
    return result.response.text();
  };
} else if (provider === "openai") {
  const OpenAI = require("openai");
  const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  _generate = async (prompt) => {
    const result = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      max_tokens: 4096,
      messages: [{ role: "user", content: prompt }],
    });
    return result.choices[0].message.content;
  };
} else {
  console.warn(
    `[aiProvider] AI_PROVIDER is "${process.env.AI_PROVIDER}" — expected "google" or "openai". AI features will be disabled.`
  );
}

console.log(`[aiProvider] Active provider: ${provider || "none"}`);

/**
 * Send a prompt to the configured AI provider and return the response text.
 * Throws if AI_PROVIDER is not set to a recognised value.
 */
async function generateText(prompt) {
  if (!_generate) {
    throw new Error(
      `AI_PROVIDER is not configured correctly. Set AI_PROVIDER=google or AI_PROVIDER=openai in .env`
    );
  }
  return _generate(prompt);
}

/** Returns the active provider name: "google" | "openai" | "" */
function getProvider() {
  return provider;
}

module.exports = { generateText, getProvider };
