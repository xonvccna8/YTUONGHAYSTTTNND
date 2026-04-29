import { GoogleGenAI } from "@google/genai";
import { getErrorMessage, readJson, setSseHeaders, validatePasscode, writeDone, writeSse } from "./_shared";

export const config = {
  maxDuration: 300,
};

export default async function handler(req: any, res: any) {
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

    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      writeSse(res, { error: "GEMINI_API_KEY chưa được cấu hình trên Vercel." });
      writeDone(res);
      return res.end();
    }

    const genAI = new GoogleGenAI({ apiKey: key });

    try {
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

      writeDone(res);
    } catch (error) {
      console.error("Gemini API Stream Error:", error);
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
