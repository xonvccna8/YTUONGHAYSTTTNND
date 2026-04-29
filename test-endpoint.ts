import fetch from "node-fetch";

async function test() {
  const response = await fetch("http://localhost:3000/api/generate-gemini", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      prompt: "Hello",
      passcode: "ideagpt2026",
      deviceId: "test-device"
    })
  });
  console.log("Status:", response.status);
  const text = await response.text();
  console.log("Body:", text);
}

test();
