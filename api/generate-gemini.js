import { GoogleGenAI } from "@google/genai";
import OpenAI from "openai";
import { getErrorMessage, readJson, setSseHeaders, validatePasscode, writeDone, writeSse } from "../lib/vercel-api.js";

export const config = {
  maxDuration: 300,
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { prompt, passcode, deviceId } = await readJson(req);

    if (!prompt || typeof prompt !== "string" || prompt.trim().length === 0) {
      return res.status(400).json({ error: "Prompt không hợp lệ hoặc bị thiếu." });
    }

    if (!passcode || !deviceId) {
      return res.status(401).json({ error: "Thiếu thông tin xác thực. Vui lòng tải lại trang." });
    }

    const authResult = validatePasscode(passcode);
    if (!authResult.valid) {
      return res.status(403).json({ error: authResult.message });
    }

    setSseHeaders(res);

    const geminiKey = process.env.GEMINI_API_KEY;
    const openaiKey = process.env.OPENAI_API_KEY;
    if (!geminiKey && !openaiKey) {
      writeSse(res, { error: "Chưa cấu hình OPENAI_API_KEY hoặc GEMINI_API_KEY trên Vercel." });
      writeDone(res);
      return res.end();
    }

    const streamWithOpenAI = async () => {
      if (!openaiKey) {
        throw new Error("OPENAI_API_KEY chưa được cấu hình trên Vercel.");
      }

      const openai = new OpenAI({ apiKey: openaiKey });
      const stream = await openai.chat.completions.create({
        model: "gpt-5.4",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.85,
        presence_penalty: 0.5,
        frequency_penalty: 0.3,
        max_completion_tokens: 4096,
        stream: true,
      });

      for await (const chunk of stream) {
        const content = chunk.choices[0]?.delta?.content || "";
        if (content) {
          writeSse(res, { text: content });
        }
      }
    };

    try {
      if (!geminiKey) {
        await streamWithOpenAI();
        writeDone(res);
        return res.end();
      }

      try {
        const genAI = new GoogleGenAI({ apiKey: geminiKey });
        const stream = await genAI.models.generateContentStream({
          model: "gemini-3.1-pro-preview",
          contents: prompt,
          config: {
            tools: [{ googleSearch: {} }],
          },
        });

        for await (const chunk of stream) {
          if (chunk.text) {
            writeSse(res, { text: chunk.text });
          }
        }
      } catch (geminiError) {
        console.error("Gemini API Stream Error:", geminiError);
        if (!openaiKey) {
          throw geminiError;
        }
        await streamWithOpenAI();
      }

      writeDone(res);
    } catch (error) {
      console.error("AI stream fallback error:", error);
      let errorMessage = getErrorMessage(error, "Lỗi không xác định từ Gemini API");

      if (errorMessage.includes("API key not valid")) {
        errorMessage = "API Key không hợp lệ hoặc đã hết hạn. Vui lòng kiểm tra cấu hình GEMINI_API_KEY trên Vercel.";
      } else if (errorMessage.includes("high demand") || errorMessage.includes("overloaded")) {
        errorMessage = "Hệ thống Gemini hiện đang quá tải. Vui lòng thử lại sau ít phút.";
      } else if (errorMessage.includes("spending cap")) {
        errorMessage = "Project Gemini đã vượt monthly spending cap. Vui lòng tăng hạn mức hoặc đổi GEMINI_API_KEY.";
      }

      writeSse(res, { error: errorMessage });
    }

    return res.end();
  } catch (error) {
    console.error("Gemini endpoint error:", error);
    return res.status(500).json({ error: getErrorMessage(error, "Lỗi không xác định") });
  }
}
