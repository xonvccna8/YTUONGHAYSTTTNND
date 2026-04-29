import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import OpenAI from "openai";
import fs from "fs";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";

dotenv.config({ override: true });

const REGISTRY_FILE = path.join(process.cwd(), "device_registry.json");
// Tạo danh sách 50 mã truy cập từ user01 đến user50. Mỗi mã chỉ dùng được trên 1 thiết bị.
const VALID_PASSCODES = Array.from({ length: 50 }, (_, i) => `user${String(i + 1).padStart(2, '0')}`);
// Giữ lại 2 mã cũ để bạn dễ dàng test
VALID_PASSCODES.push("ideagpt2026", "vip2026");

const UNLIMITED_TIME = 100 * 365 * 24 * 60 * 60 * 1000; // 100 years

type PasscodeCheckResult = {
  valid: boolean;
  message?: string;
  isExpired?: boolean;
  expiresAt?: number;
  registry?: Record<string, any>;
};

function checkPasscodeValid(passcode: string, deviceId: string): PasscodeCheckResult {
  if (!VALID_PASSCODES.includes(passcode)) {
    return { valid: false, message: "Mã truy cập không tồn tại hoặc không hợp lệ." };
  }

  let registry: Record<string, any> = {};
  if (fs.existsSync(REGISTRY_FILE)) {
    try {
      registry = JSON.parse(fs.readFileSync(REGISTRY_FILE, "utf-8"));
    } catch (e) {
      console.error("Error reading registry file", e);
    }
  }

  const entry = registry[passcode];
  if (entry) {
    // Hỗ trợ cả định dạng cũ (string) và mới (object)
    const regDeviceId = typeof entry === 'string' ? entry : entry.deviceId;
    const firstLoginAt = typeof entry === 'string' ? Date.now() : entry.firstLoginAt;

    if (regDeviceId !== deviceId) {
      return { valid: false, message: "Mã truy cập này đã được sử dụng trên một thiết bị khác!" };
    }

    // Đã bỏ giới hạn 30 phút
    // if (Date.now() - firstLoginAt > THIRTY_MINUTES) { ... }

    return { valid: true, expiresAt: firstLoginAt + UNLIMITED_TIME, registry };
  } else {
    // Đăng nhập lần đầu tiên cho mã này
    registry[passcode] = { deviceId, firstLoginAt: Date.now() };
    try {
      fs.writeFileSync(REGISTRY_FILE, JSON.stringify(registry, null, 2));
    } catch (e) {
      console.error("Error writing registry file", e);
    }
    return { valid: true, expiresAt: Date.now() + UNLIMITED_TIME, registry };
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  function getOpenAI(): OpenAI {
    const key = process.env.OPENAI_API_KEY;
    if (!key) {
      throw new Error("OPENAI_API_KEY environment variable is required");
    }
    return new OpenAI({ apiKey: key });
  }

  function getDeepSeek(): OpenAI {
    const key = process.env.DEEPSEEK_API_KEY;
    if (!key) {
      throw new Error("DEEPSEEK_API_KEY environment variable is required");
    }
    return new OpenAI({ apiKey: key, baseURL: "https://api.deepseek.com" });
  }

  // Endpoint xác thực mã truy cập và khóa thiết bị
  app.post("/api/verify-passcode", (req, res) => {
    const { passcode, deviceId } = req.body;

    if (!passcode || !deviceId) {
      return res.status(400).json({ success: false, message: "Thiếu thông tin xác thực." });
    }

    const result = checkPasscodeValid(passcode, deviceId);
    
    if (!result.valid) {
      return res.status(403).json({ 
        success: false, 
        isExpired: result.isExpired, 
        message: result.message 
      });
    }

    return res.json({ 
      success: true, 
      message: "Xác thực thành công.", 
      expiresAt: result.expiresAt 
    });
  });

  app.post("/api/generate", async (req, res) => {
    const keepAliveInterval = null as any;
    
    try {
      const { prompt, passcode, deviceId, mode } = req.body;

      // Validate input
      if (!prompt || typeof prompt !== 'string' || prompt.trim().length === 0) {
        return res.status(400).json({ error: "Prompt không hợp lệ hoặc bị thiếu." });
      }

      // Kiểm tra bảo mật trước khi generate
      if (!passcode || !deviceId) {
        return res.status(401).json({ error: "Thiếu thông tin xác thực. Vui lòng tải lại trang." });
      }

      const authResult = checkPasscodeValid(passcode, deviceId);
      if (!authResult.valid) {
        return res.status(403).json({ 
          error: authResult.message, 
          isExpired: authResult.isExpired 
        });
      }

      // Select model based on mode
      const modelMap: Record<string, string> = {
        "advanced-gpt": "gpt-5.4",
        "basic-gpt": "gpt-5.4-mini"
      };
      const model = modelMap[mode] || "gpt-5.4-mini";

      // Setup SSE headers
      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");
      res.setHeader("X-Accel-Buffering", "no"); // Disable nginx buffering

      // Keep connection alive with ping
      const keepAliveInterval = setInterval(() => {
        if (!res.writableEnded) {
          res.write(': ping\n\n');
        }
      }, 15000);

      // Cleanup on client disconnect
      req.on('close', () => {
        clearInterval(keepAliveInterval);
      });

      try {
        const openai = getOpenAI();
        const stream = await openai.chat.completions.create({
          model: model,
          messages: [{ role: "user", content: prompt }],
          temperature: 0.85,
          presence_penalty: 0.5,
          frequency_penalty: 0.3,
          max_completion_tokens: 16384,
          stream: true,
        });

        for await (const chunk of stream) {
          if (res.writableEnded) break;
          
          const content = chunk.choices[0]?.delta?.content || "";
          if (content) {
            res.write(`data: ${JSON.stringify({ text: content })}\n\n`);
          }

          // Check for finish reason
          const finishReason = chunk.choices[0]?.finish_reason;
          if (finishReason) {
            console.log(`Stream finished with reason: ${finishReason}`);
          }
        }

        if (!res.writableEnded) {
          res.write("data: [DONE]\n\n");
        }
      } catch (streamError: any) {
        console.error("OpenAI API Stream Error:", streamError);
        
        let errorMessage = "Lỗi không xác định từ OpenAI API";
        
        // Handle specific OpenAI errors
        if (streamError.status === 401) {
          errorMessage = "API Key không hợp lệ. Vui lòng kiểm tra cấu hình OPENAI_API_KEY.";
        } else if (streamError.status === 429) {
          errorMessage = "Đã vượt quá giới hạn request. Vui lòng thử lại sau.";
        } else if (streamError.status === 500) {
          errorMessage = "Lỗi server từ OpenAI. Vui lòng thử lại sau.";
        } else if (streamError.status === 503) {
          errorMessage = "Dịch vụ OpenAI tạm thời không khả dụng. Vui lòng thử lại sau.";
        } else if (streamError.message) {
          errorMessage = streamError.message;
        }

        if (!res.writableEnded) {
          res.write(`data: ${JSON.stringify({ error: errorMessage })}\n\n`);
        }
      } finally {
        clearInterval(keepAliveInterval);
        if (!res.writableEnded) {
          res.end();
        }
      }
    } catch (error: any) {
      console.error("OpenAI API Error:", error);
      
      // Clear interval if it exists
      if (keepAliveInterval) {
        clearInterval(keepAliveInterval);
      }

      if (!res.headersSent) {
        res.status(500).json({ 
          error: error.message || "Lỗi không xác định khi xử lý request" 
        });
      } else if (!res.writableEnded) {
        res.end();
      }
    }
  });

  app.post("/api/generate-deepseek", async (req, res) => {
    try {
      const { prompt, passcode, deviceId } = req.body;

      if (!prompt || typeof prompt !== 'string' || prompt.trim().length === 0) {
        return res.status(400).json({ error: "Prompt không hợp lệ hoặc bị thiếu." });
      }

      if (!passcode || !deviceId) {
        return res.status(401).json({ error: "Thiếu thông tin xác thực. Vui lòng tải lại trang." });
      }

      const authResult = checkPasscodeValid(passcode, deviceId);
      if (!authResult.valid) {
        return res.status(403).json({
          error: authResult.message,
          isExpired: authResult.isExpired
        });
      }

      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");
      res.setHeader("X-Accel-Buffering", "no");

      const keepAliveInterval = setInterval(() => {
        if (!res.writableEnded) {
          res.write(': ping\n\n');
        }
      }, 15000);

      req.on('close', () => {
        clearInterval(keepAliveInterval);
      });

      try {
        const deepseek = getDeepSeek();
        const stream = await deepseek.chat.completions.create({
          model: "deepseek-v4-pro",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.85,
          presence_penalty: 0.5,
          frequency_penalty: 0.3,
          thinking: { type: "disabled" },
          max_tokens: 65536,
          stream: true,
        } as any) as any;

        for await (const chunk of stream) {
          if (res.writableEnded) break;
          const content = chunk.choices[0]?.delta?.content || "";
          if (content) {
            res.write(`data: ${JSON.stringify({ text: content })}\n\n`);
          }
        }

        if (!res.writableEnded) {
          res.write("data: [DONE]\n\n");
        }
      } catch (streamError: any) {
        console.error("DeepSeek API Stream Error:", streamError);

        let errorMessage = "Lỗi không xác định từ DeepSeek API";
        if (streamError.status === 401) {
          errorMessage = "DeepSeek API Key không hợp lệ. Vui lòng kiểm tra cấu hình DEEPSEEK_API_KEY.";
        } else if (streamError.status === 402) {
          errorMessage = "Tài khoản DeepSeek không đủ số dư. Vui lòng kiểm tra Billing DeepSeek.";
        } else if (streamError.status === 429) {
          errorMessage = "DeepSeek đang giới hạn request hoặc quá tải. Vui lòng thử lại sau.";
        } else if (streamError.status >= 500) {
          errorMessage = "Dịch vụ DeepSeek tạm thời không khả dụng. Vui lòng thử lại sau.";
        } else if (streamError.message) {
          errorMessage = streamError.message;
        }

        if (!res.writableEnded) {
          res.write(`data: ${JSON.stringify({ error: errorMessage })}\n\n`);
        }
      } finally {
        clearInterval(keepAliveInterval);
        if (!res.writableEnded) {
          res.end();
        }
      }
    } catch (error: any) {
      console.error("DeepSeek API Error:", error);
      if (!res.headersSent) {
        res.status(500).json({ error: error.message || "Lỗi không xác định" });
      } else if (!res.writableEnded) {
        res.end();
      }
    }
  });

  function getGenAI(): GoogleGenAI {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error("GEMINI_API_KEY environment variable is required");
    }
    return new GoogleGenAI({ apiKey: key });
  }

  app.post("/api/generate-gemini", async (req, res) => {
    try {
      const { prompt, passcode, deviceId } = req.body;

      if (!passcode || !deviceId) {
        return res.status(401).json({ error: "Thiếu thông tin xác thực. Vui lòng tải lại trang." });
      }

      const authResult = checkPasscodeValid(passcode, deviceId);
      if (!authResult.valid) {
        return res.status(403).json({ 
          error: authResult.message, 
          isExpired: authResult.isExpired 
        });
      }

      const genAI = getGenAI();

      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");

      // Send a ping every 15 seconds to keep the connection alive
      const keepAliveInterval = setInterval(() => {
        if (!res.writableEnded) {
          res.write(': ping\n\n');
        }
      }, 15000);

      req.on('close', () => {
        clearInterval(keepAliveInterval);
      });

      try {
        const stream = await genAI.models.generateContentStream({
          model: "gemini-3.1-pro-preview",
          contents: prompt,
          config: {
            tools: [{ googleSearch: {} }]
          }
        });

        for await (const chunk of stream) {
          if (res.writableEnded) break;
          if (chunk.text) {
            res.write(`data: ${JSON.stringify({ text: chunk.text })}\n\n`);
          }
        }
        if (!res.writableEnded) {
          res.write("data: [DONE]\n\n");
        }
      } catch (streamError: any) {
        console.error("Gemini API Stream Error:", streamError);
        let errorMessage = "Lỗi không xác định từ Gemini API";
        
        if (streamError instanceof Error) {
          errorMessage = streamError.message;
          try {
            const parsedError = JSON.parse(streamError.message);
            if (parsedError.error && parsedError.error.message) {
              const innerError = JSON.parse(parsedError.error.message);
              if (innerError.error && innerError.error.message) {
                errorMessage = innerError.error.message;
              }
            }
          } catch (e) {
            // Ignore parse error
          }
        } else {
          errorMessage = String(streamError);
        }

        if (errorMessage.includes("API key not valid")) {
          errorMessage = "API Key không hợp lệ hoặc đã hết hạn. Vui lòng kiểm tra lại cấu hình GEMINI_API_KEY trong phần Secrets của AI Studio.";
        } else if (errorMessage.includes("high demand") || errorMessage.includes("overloaded")) {
          errorMessage = "Hệ thống Gemini hiện đang quá tải do nhu cầu sử dụng cao. Vui lòng thử lại sau ít phút.";
        }

        if (!res.writableEnded) {
          res.write(`data: ${JSON.stringify({ error: errorMessage })}\n\n`);
        }
      } finally {
        clearInterval(keepAliveInterval);
        if (!res.writableEnded) {
          res.end();
        }
      }
    } catch (error: any) {
      console.error("Gemini API Error:", error);
      if (!res.headersSent) {
        res.status(500).json({ error: error.message || "Lỗi không xác định" });
      } else if (!res.writableEnded) {
        res.end();
      }
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
