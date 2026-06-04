const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.post('/optimize', async (req, res) => {
  const { resume, jobDescription } = req.body;
  if (!resume || !jobDescription) {
    return res.status(400).json({ error: 'Missing resume or job description' });
  }
  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4000,
      messages: [{
        role: 'user',
        content: `You are an expert resume coach and ATS optimization specialist.

Given this resume and job description, provide:
1. An ATS SCORE (0-100) for the original resume
2. KEY MISSING KEYWORDS from the job description
3. AN OPTIMIZED RESUME tailored to the job
4. A COVER LETTER for this specific role

Format your response exactly like this:
---ATS SCORE---
[score]/100 - [one line explanation]

---MISSING KEYWORDS---
[comma separated keywords]

---OPTIMIZED RESUME---
[full optimized resume]

---COVER LETTER---
[full cover letter]

RESUME:
${resume}

JOB DESCRIPTION:
${jobDescription}`
      }]
    });
    res.json({ result: message.content[0].text });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(3000, () => console.log('ResumeAI running on http://localhost:3000'));