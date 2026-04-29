import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
dotenv.config({ override: true });

async function test() {
  try {
    const genAI = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
    const stream = await genAI.models.generateContentStream({
      model: "gemini-3.1-pro-preview",
      contents: "Hello",
    });
    for await (const chunk of stream) {
      console.log(chunk.text);
    }
  } catch (e) {
    console.error("Caught error:", e);
  }
}
test();
