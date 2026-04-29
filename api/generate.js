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
    const { prompt, passcode, deviceId, mode } = await readJson(req);

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

    const key = process.env.OPENAI_API_KEY;
    if (!key) {
      writeSse(res, { error: "OPENAI_API_KEY chưa được cấu hình trên Vercel." });
      writeDone(res);
      return res.end();
    }

    const modelMap = {
      "advanced-gpt": "gpt-5.4",
      "basic-gpt": "gpt-5.4-mini",
    };
    const model = modelMap[mode] || "gpt-5.4-mini";
    const openai = new OpenAI({ apiKey: key });

    try {
      const stream = await openai.chat.completions.create({
        model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.85,
        presence_penalty: 0.5,
        frequency_penalty: 0.3,
        max_completion_tokens: 16384,
        stream: true,
      });

      for await (const chunk of stream) {
        const content = chunk.choices[0]?.delta?.content || "";
        if (content) {
          writeSse(res, { text: content });
        }
      }

      writeDone(res);
    } catch (error) {
      console.error("OpenAI API Stream Error:", error);

      let errorMessage = getErrorMessage(error, "Lỗi không xác định từ OpenAI API");
      if (error.status === 401) {
        errorMessage = "API Key không hợp lệ. Vui lòng kiểm tra cấu hình OPENAI_API_KEY.";
      } else if (error.status === 429) {
        errorMessage = "Đã vượt quá giới hạn request hoặc quota. Vui lòng thử lại sau.";
      } else if (error.status === 500) {
        errorMessage = "Lỗi server từ OpenAI. Vui lòng thử lại sau.";
      } else if (error.status === 503) {
        errorMessage = "Dịch vụ OpenAI tạm thời không khả dụng. Vui lòng thử lại sau.";
      }

      writeSse(res, { error: errorMessage });
    }

    return res.end();
  } catch (error) {
    console.error("Generate endpoint error:", error);
    return res.status(500).json({ error: getErrorMessage(error, "Lỗi không xác định khi xử lý request") });
  }
}
