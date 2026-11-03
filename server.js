// server.js
import dotenv from "dotenv";
dotenv.config();

import express from "express";
import pptxgen from "pptxgenjs";
import fs from "fs";
import path from "path";
import fetch from "node-fetch";

const app = express();
const PORT = 7000;

// Middleware
app.use(express.json());
app.use(express.static(path.join(process.cwd(), "public")));

// Check API key
const API_KEY = process.env.OPENAI_API_KEY;
if (!API_KEY) {
  console.error("❌ OPENAI_API_KEY not set in .env");
  process.exit(1);
}

// ================= GENERATE PPT =================
app.post("/generate", async (req, res) => {
  try {
    const { topic, slides, audience, style } = req.body;

    if (!topic || !slides) {
      return res.status(400).json({ error: "Topic and slides are required" });
    }

    // -------- OpenAI API CALL --------
    const aiResponse = await fetch(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${API_KEY}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "gpt-4o",
          messages: [
            {
              role: "system",
              content:
                "You are an expert presentation designer. Create slide-wise bullet content.",
            },
            {
              role: "user",
              content: `
Create ${slides} slides on "${topic}"
Audience: ${audience}
Style: ${style}

Format EXACTLY like:
Slide 1: Title
- Bullet
- Bullet
Slide 2: Title
- Bullet
`,
            },
          ],
          temperature: 0.7,
        }),
      }
    );

    if (!aiResponse.ok) {
      const err = await aiResponse.text();
      console.error("❌ OpenAI Error:", err);
      return res.status(500).json({ error: "OpenAI failed" });
    }

    const aiData = await aiResponse.json();
    const content = aiData.choices[0].message.content;

    // -------- CREATE PPT --------
    const pptx = new pptxgen();
    pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
    pptx.layout = "WIDE";

    // Title slide
    const titleSlide = pptx.addSlide();
    titleSlide.addText(topic, {
      x: 1,
      y: 2,
      w: 11,
      fontSize: 36,
      bold: true,
      align: "center",
    });

    // Content slides
    const slideBlocks = content.split(/Slide \d+:/).filter(Boolean);

    slideBlocks.forEach((block) => {
      const lines = block.trim().split("\n").filter(Boolean);
      const title = lines[0];
      const bullets = lines
        .slice(1)
        .map((l) => l.replace(/^-/, "").trim());

      const slide = pptx.addSlide();

      slide.addText(title, {
        x: 0.5,
        y: 0.5,
        fontSize: 26,
        bold: true,
      });

      slide.addText(bullets, {
        x: 0.7,
        y: 1.5,
        w: 12,
        h: 5,
        fontSize: 18,
        bullet: true,
      });
    });

    // Save file
    const fileName = `AI_Presentation_${Date.now()}.pptx`;
    const filePath = path.join(process.cwd(), fileName);
    await pptx.writeFile({ fileName: filePath });

    // Send to frontend
    res.download(filePath, fileName, () => {
      fs.unlinkSync(filePath);
    });
  } catch (err) {
    console.error("❌ SERVER ERROR:", err);
    res.status(500).json({ error: "PPT generation failed" });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
