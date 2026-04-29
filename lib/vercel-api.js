export const VALID_PASSCODES = [
  ...Array.from({ length: 50 }, (_, index) => `user${String(index + 1).padStart(2, "0")}`),
  "ideagpt2026",
  "vip2026",
];

export const UNLIMITED_TIME = 100 * 365 * 24 * 60 * 60 * 1000;

export async function readJson(req) {
  if (req.body && typeof req.body === "object") {
    return req.body;
  }

  if (typeof req.body === "string") {
    return JSON.parse(req.body || "{}");
  }

  let raw = "";
  await new Promise((resolve, reject) => {
    req.on("data", chunk => {
      raw += chunk;
    });
    req.on("end", resolve);
    req.on("error", reject);
  });

  return raw ? JSON.parse(raw) : {};
}

export function validatePasscode(passcode) {
  if (typeof passcode !== "string" || !VALID_PASSCODES.includes(passcode)) {
    return { valid: false, message: "Mã truy cập không tồn tại hoặc không hợp lệ." };
  }

  return {
    valid: true,
    expiresAt: Date.now() + UNLIMITED_TIME,
    message: "Xác thực thành công.",
  };
}

export function setSseHeaders(res) {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
}

export function writeSse(res, payload) {
  res.write(`data: ${JSON.stringify(payload)}\n\n`);
}

export function writeDone(res) {
  res.write("data: [DONE]\n\n");
}

export function getErrorMessage(error, fallback = "Lỗi không xác định") {
  return error instanceof Error ? error.message : String(error || fallback);
}
