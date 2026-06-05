const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const Stripe = require('stripe');
const { parseResumeToDocx, parseCoverLetterToDocx } = require('./generate_docx');
require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(cors());
app.use('/webhook', express.raw({ type: 'application/json' }));
app.use(express.json({ limit: '10mb' }));

// Routes BEFORE static
app.get('/', (req, res) => res.sendFile(__dirname + '/public/landing.html'));
app.get('/app', (req, res) => res.sendFile(__dirname + '/public/index.html'));

app.use(express.static('public'));

// ── SYSTEMS ──────────────────────────────────────────────────────────────────

const CV_FORMAT = `
CRITICAL: Output the CV as plain text using this EXACT structure. Use 3+ spaces to separate entry names from dates.

[Full Name]
[City, Country] • [Phone] • [Email]

EDUCATION

[University Name]   [City], [Month Year - Month Year]
Course: [Degree name]
Final qualification: [Full degree title]
GPA/Grade: [if strong]

WORK EXPERIENCE

[Company Name - Role Title]   [City], [Month Year - Month Year]
- [Action verb + specific task + quantified result]
- [Action verb + specific task + quantified result]
- [Action verb + specific task + quantified result]

EXTRACURRICULAR ACTIVITIES

[Activity Name - Role]   [City], [Month Year - Month Year]
- [Bullet point]

SKILLS AND INTERESTS

Languages: [Language (Level)]
Skills: [Skill 1, Skill 2, Skill 3]
Interests: [Interest 1, Interest 2, Interest 3]

RULES:
- ONE PAGE maximum
- Start every bullet with a strong action verb (Executed, Analysed, Built, Led, Generated, Managed)
- Quantify everything — numbers, percentages, £/€ amounts
- No asterisks, no markdown, no emojis
- 3+ spaces between entry name and date
`;

const STUDENT_SYSTEM = `You are an expert career coach for student applications — internships, spring weeks, graduate schemes at top finance, consulting and tech firms. You maximise limited experience by highlighting extracurriculars, academics and transferable skills. You write in a confident, achievement-focused style. ${CV_FORMAT}`;

const PROFESSIONAL_SYSTEM = `You are an expert executive resume writer placing candidates at Fortune 500 companies. You quantify achievements, position career pivots, and craft ATS-optimised narratives. You write in a results-driven, metrics-focused style. ${CV_FORMAT}`;

const TARGETED_SYSTEM = `You are an elite career strategist building highly targeted applications for specific companies and programs. You have deep knowledge of investment banking, consulting, and tech hiring. You know exactly what JP Morgan, Goldman Sachs, McKinsey, Google and hundreds of other firms look for. Always use your web search to find current, specific information about the program before writing. Your applications are indistinguishable from those written by insiders.`;

const buildOptimisePrompt = (resume, jobDesc, userType, templateStyle) => `
Analyse this resume against the job description. User type: ${userType === 'student' ? 'Student (internship/spring week/grad scheme)' : 'Professional (full-time role)'}.

Respond in EXACTLY this format with these exact markers:

---ATS SCORE---
[number]/100 - [one sentence explanation]

---MISSING KEYWORDS---
[comma separated keywords from job description missing in resume]

---OPTIMIZED RESUME---
[Full optimised CV following the format rules exactly]

---COVER LETTER---
[Tailored cover letter, 3-4 paragraphs, professional tone]

---INTERVIEW TIPS---
- [Specific tip 1 for this role]
- [Specific tip 2 for this role]  
- [Specific tip 3 for this role]

---FEEDBACK CARDS---
[JSON array of 3 improvement suggestions. Each: {"category":"hobbies|experience|skills|education|formatting|keywords","issue":"one sentence problem","options":["suggestion 1","suggestion 2","suggestion 3"]}]

RESUME:
${resume}

JOB DESCRIPTION:
${jobDesc}
`;

const buildTargetedPrompt = (resume, target) => `
The user wants to apply to: "${target}"

Search the web to find:
1. Exact details about this program/role (requirements, deadlines, what they look for)
2. The company's current values, culture, recent news
3. What makes a STRONG application for this specific program
4. Typical interview process and assessment criteria

Then build a complete application package using the research AND the user's resume.

Respond in EXACTLY this format:

---RESEARCH SUMMARY---
[3-4 specific sentences about what you found — include real details about the program]

---APPLICATION STRATEGY---
[2-3 sentences on the specific angle to take based on research]

---TAILORED CV---
[Full CV rewritten specifically for this company, using their language and values, following CV format rules]

---COVER LETTER---
[Cover letter specifically referencing this company's programs, values, recent news — 3-4 paragraphs]

---INTERVIEW PREP---
- [Likely question 1 + brief answer framework]
- [Likely question 2 + brief answer framework]
- [Likely question 3 + brief answer framework]
- [Likely question 4 + brief answer framework]
- [Likely question 5 + brief answer framework]

USER'S RESUME:
${resume}
`;

// ── OPTIMISE ─────────────────────────────────────────────────────────────────

app.post('/optimize', async (req, res) => {
  const { resume, jobDescription, userType, templateStyle } = req.body;
  if (!resume || !jobDescription) return res.status(400).json({ error: 'Missing fields' });
  try {
    const msg = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4000,
      system: userType === 'student' ? STUDENT_SYSTEM : PROFESSIONAL_SYSTEM,
      messages: [{ role: 'user', content: buildOptimisePrompt(resume, jobDescription, userType, templateStyle) }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── TARGETED ─────────────────────────────────────────────────────────────────

app.post('/targeted', async (req, res) => {
  const { resume, target } = req.body;
  if (!resume || !target) return res.status(400).json({ error: 'Missing fields' });

  try {
    // Full agentic loop for web search
    const messages = [{ role: 'user', content: buildTargetedPrompt(resume, target) }];

    let response = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 8000,
      system: TARGETED_SYSTEM,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }],
      messages
    });

    // Handle tool_use loops (max 8 iterations)
    let iter = 0;
    while (response.stop_reason === 'tool_use' && iter < 8) {
      iter++;
      // Append assistant turn
      messages.push({ role: 'assistant', content: response.content });

      // Build tool results
      const toolResults = response.content
        .filter(b => b.type === 'tool_use')
        .map(b => ({
          type: 'tool_result',
          tool_use_id: b.id,
          content: b.type === 'tool_use' ? `Search for "${JSON.stringify(b.input)}" completed.` : ''
        }));

      messages.push({ role: 'user', content: toolResults });

      response = await client.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 8000,
        system: TARGETED_SYSTEM,
        tools: [{ type: 'web_search_20250305', name: 'web_search' }],
        messages
      });
    }

    // Extract all text
    let fullText = response.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n');

    console.log(`Targeted done. Iterations: ${iter}, Text: ${fullText.length} chars`);

    // Fallback: if text is too short, run without web search
    if (fullText.trim().length < 200) {
      console.log('Web search produced no content, using knowledge fallback');
      const fallback = await client.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 8000,
        system: TARGETED_SYSTEM,
        messages: [{ role: 'user', content: buildTargetedPrompt(resume, target) }]
      });
      fullText = fallback.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
    }

    if (!fullText.trim()) return res.status(500).json({ error: 'No content generated. Please try again.' });
    res.json({ result: fullText.trim() });

  } catch (e) {
    console.error('Targeted error:', e.message);
    // Final fallback without tools
    try {
      const fb = await client.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 8000,
        system: TARGETED_SYSTEM,
        messages: [{ role: 'user', content: buildTargetedPrompt(resume, target) }]
      });
      const text = fb.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
      res.json({ result: text.trim() });
    } catch (e2) {
      res.status(500).json({ error: e.message });
    }
  }
});

// ── DOCX DOWNLOADS ───────────────────────────────────────────────────────────

app.post('/download-cv', async (req, res) => {
  const { resumeText } = req.body;
  if (!resumeText) return res.status(400).json({ error: 'Missing resume text' });
  try {
    const buf = await parseResumeToDocx(resumeText);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="optimised-cv.docx"');
    res.send(buf);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/download-cover', async (req, res) => {
  const { coverText } = req.body;
  if (!coverText) return res.status(400).json({ error: 'Missing cover letter text' });
  try {
    const buf = await parseCoverLetterToDocx(coverText);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="cover-letter.docx"');
    res.send(buf);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── TEMPLATE ANALYSER ────────────────────────────────────────────────────────

app.post('/analyze-template', upload.single('template'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  try {
    const txt = fs.readFileSync(req.file.path, 'utf8');
    fs.unlinkSync(req.file.path);
    const msg = await client.messages.create({
      model: 'claude-sonnet-4-6', max_tokens: 300,
      messages: [{ role: 'user', content: `Describe this CV template style in 2 sentences:\n${txt}` }]
    });
    res.json({ style: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── STRIPE ───────────────────────────────────────────────────────────────────

app.post('/create-checkout', async (req, res) => {
  const { userId, userEmail } = req.body;
  if (!userId || !userEmail) return res.status(400).json({ error: 'Missing user info' });
  try {
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      mode: 'subscription',
      customer_email: userEmail,
      line_items: [{ price: process.env.STRIPE_PRICE_ID, quantity: 1 }],
      metadata: { userId },
      success_url: `${process.env.APP_URL}/app?upgrade=success`,
      cancel_url: `${process.env.APP_URL}/app?upgrade=cancelled`,
    });
    res.json({ url: session.url });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/check-subscription', async (req, res) => {
  const { userEmail } = req.body;
  if (!userEmail) return res.json({ isPro: false });
  try {
    const customers = await stripe.customers.list({ email: userEmail, limit: 1 });
    if (!customers.data.length) return res.json({ isPro: false });
    const subs = await stripe.subscriptions.list({ customer: customers.data[0].id, status: 'active', limit: 1 });
    res.json({ isPro: subs.data.length > 0 });
  } catch (e) { res.json({ isPro: false }); }
});

app.post('/webhook', async (req, res) => {
  const sig = req.headers['stripe-signature'];
  try {
    stripe.webhooks.constructEvent(req.body, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (e) { return res.status(400).send(`Webhook Error: ${e.message}`); }
  res.json({ received: true });
});

// ── START ─────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ResumeAI running on port ${PORT}`));
