import { readJson, validatePasscode } from "./_shared";

export default async function handler(req: any, res: any) {
  if (req.method !== "POST") {
    return res.status(405).json({ success: false, message: "Method not allowed" });
  }

  try {
    const { passcode, deviceId } = await readJson(req);

    if (!passcode || !deviceId) {
      return res.status(400).json({ success: false, message: "Thiếu thông tin xác thực." });
    }

    const result = validatePasscode(passcode);
    if (!result.valid) {
      return res.status(403).json({ success: false, message: result.message });
    }

    return res.json({
      success: true,
      message: result.message,
      expiresAt: result.expiresAt,
    });
  } catch (error) {
    console.error("Verify passcode error:", error);
    return res.status(500).json({ success: false, message: "Lỗi máy chủ. Vui lòng thử lại." });
  }
}
