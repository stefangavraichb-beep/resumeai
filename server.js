const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const Stripe = require('stripe');
const { buildFromElements, parseCoverLetterToDocx } = require('./generate_docx');

require('dotenv').config();

const rateLimits = new Map();
function rateLimit(maxReqs, windowMs) {
  return (req, res, next) => {
    const key = req.ip || 'unknown';
    const now = Date.now();
    const entry = rateLimits.get(key) || { count: 0, resetAt: now + windowMs };
    if (now > entry.resetAt) { entry.count = 0; entry.resetAt = now + windowMs; }
    entry.count++;
    rateLimits.set(key, entry);
    if (entry.count > maxReqs) return res.status(429).json({ error: 'Too many requests. Please wait a moment.' });
    next();
  };
}

const app = express();
const upload = multer({ dest: 'uploads/' });
const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(cors());
app.use('/webhook', express.raw({ type: 'application/json' }));
app.use(express.json({ limit: '10mb' }));

// Security headers
app.use((req, res, next) => {
  res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains; preload');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
  next();
});

// WWW redirect
app.use((req, res, next) => {
  if (req.headers.host === 'www.resumeai.today') {
    return res.redirect(301, 'https://resumeai.today' + req.url);
  }
  next();
});

// SEO
app.get('/robots.txt', (req, res) => {
  res.type('text/plain');
  res.send('User-agent: *\nAllow: /\nDisallow: /app\nDisallow: /auth.html\n\nSitemap: https://resumeai.today/sitemap.xml');
});

app.get('/sitemap.xml', (req, res) => {
  res.type('application/xml');
  res.send('<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"><url><loc>https://resumeai.today/</loc><changefreq>weekly</changefreq><priority>1.0</priority></url></urlset>');
});

// Page routes
app.get('/', (req, res) => res.sendFile(__dirname + '/public/landing.html'));
app.get('/app', (req, res) => res.sendFile(__dirname + '/public/index.html'));

// Static files
app.use(express.static('public'));

// ── API ROUTES (must be BEFORE 404 handler) ───────────────────────────────────

const CV_FORMAT = `
CRITICAL: Output the CV as plain text using this EXACT structure. Use 3+ spaces to separate entry names from dates.

[Full Name]
[City, Country] • [Phone] • [Email]

EDUCATION
[University Name]   [City], [Month Year - Month Year]
Course: [Degree name]

WORK EXPERIENCE
[Company Name - Role Title]   [City], [Month Year - Month Year]
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
- Start every bullet with a strong action verb
- Quantify everything
- No asterisks, no markdown, no emojis
- 3+ spaces between entry name and date
`;

const STUDENT_SYSTEM = `You are an expert career coach for student applications — internships, spring weeks, graduate schemes at top finance, consulting and tech firms. ${CV_FORMAT}`;
const PROFESSIONAL_SYSTEM = `You are an expert executive resume writer placing candidates at Fortune 500 companies. ${CV_FORMAT}`;
const TARGETED_SYSTEM = `You are an elite career strategist building highly targeted applications for specific companies and programs. You have deep knowledge of investment banking, consulting, and tech hiring. Always use web search to find current, specific information before writing.`;

const buildOptimisePrompt = (resume, jobDesc, userType) => `
Analyse this resume against the job description. User type: ${userType === 'student' ? 'Student' : 'Professional'}.

Respond in EXACTLY this format:

---ATS SCORE---
[number]/100 - [one sentence explanation]

---MISSING KEYWORDS---
[comma separated keywords missing in resume]

---OPTIMIZED RESUME---
[Full optimised CV]

---COVER LETTER---
[Tailored cover letter, 3-4 paragraphs]

---INTERVIEW TIPS---
- [Tip 1]
- [Tip 2]
- [Tip 3]

---FEEDBACK CARDS---
[JSON array: {"category":"hobbies|experience|skills|education|formatting|keywords","issue":"problem","options":["opt1","opt2","opt3"]}]

RESUME:
${resume}

JOB DESCRIPTION:
${jobDesc}
`;

const buildTargetedPrompt = (resume, target) => `
The user wants to apply to: "${target}"

Search the web for details about this program, then build a complete application.

Respond in EXACTLY this format:

---RESEARCH SUMMARY---
[3-4 specific sentences about the program]

---APPLICATION STRATEGY---
[2-3 sentences on the angle to take]

---TAILORED CV---
[Full CV rewritten for this company]

---COVER LETTER---
[Cover letter referencing company specifics]

---INTERVIEW PREP---
- [Question 1 + answer framework]
- [Question 2 + answer framework]
- [Question 3 + answer framework]
- [Question 4 + answer framework]
- [Question 5 + answer framework]

USER'S RESUME:
${resume}
`;

app.post('/optimize', rateLimit(20, 60000), async (req, res) => {
  const { resume, jobDescription, userType } = req.body;
  if (!resume || !jobDescription) return res.status(400).json({ error: 'Missing fields' });
  try {
    const msg = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4000,
      system: userType === 'student' ? STUDENT_SYSTEM : PROFESSIONAL_SYSTEM,
      messages: [{ role: 'user', content: buildOptimisePrompt(resume, jobDescription, userType) }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/targeted', rateLimit(10, 60000), async (req, res) => {
  const { resume, target } = req.body;
  if (!resume || !target) return res.status(400).json({ error: 'Missing fields' });
  try {
    const messages = [{ role: 'user', content: buildTargetedPrompt(resume, target) }];
    let response = await client.messages.create({
      model: 'claude-sonnet-4-6', max_tokens: 8000, system: TARGETED_SYSTEM,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }], messages
    });
    let iter = 0;
    while (response.stop_reason === 'tool_use' && iter < 8) {
      iter++;
      messages.push({ role: 'assistant', content: response.content });
      const toolResults = response.content.filter(b => b.type === 'tool_use').map(b => ({ type: 'tool_result', tool_use_id: b.id, content: `Search completed for: ${JSON.stringify(b.input)}` }));
      messages.push({ role: 'user', content: toolResults });
      response = await client.messages.create({ model: 'claude-sonnet-4-6', max_tokens: 8000, system: TARGETED_SYSTEM, tools: [{ type: 'web_search_20250305', name: 'web_search' }], messages });
    }
    let fullText = response.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
    if (fullText.length < 200) {
      const fb = await client.messages.create({ model: 'claude-sonnet-4-6', max_tokens: 8000, system: TARGETED_SYSTEM, messages: [{ role: 'user', content: buildTargetedPrompt(resume, target) }] });
      fullText = fb.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
    }
    if (!fullText) return res.status(500).json({ error: 'No content generated.' });
    res.json({ result: fullText });
  } catch (e) {
    try {
      const fb = await client.messages.create({ model: 'claude-sonnet-4-6', max_tokens: 8000, system: TARGETED_SYSTEM, messages: [{ role: 'user', content: buildTargetedPrompt(resume, target) }] });
      res.json({ result: fb.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim() });
    } catch (e2) { res.status(500).json({ error: e.message }); }
  }
});

app.post('/download-cv', async (req, res) => {
  const { elements } = req.body;
  if (!elements || !elements.length) return res.status(400).json({ error: 'Missing elements' });
  try {
    const buf = await buildFromElements(elements);
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

app.post('/analyze-template', upload.single('template'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  try {
    const txt = fs.readFileSync(req.file.path, 'utf8');
    fs.unlinkSync(req.file.path);
    const msg = await client.messages.create({ model: 'claude-sonnet-4-6', max_tokens: 300, messages: [{ role: 'user', content: `Describe this CV template style in 2 sentences:\n${txt}` }] });
    res.json({ style: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/capture-email', async (req, res) => {
  const { email, userId } = req.body;
  if (!email) return res.status(400).json({ error: 'Missing email' });
  try {
    fs.appendFileSync('captured_emails.txt', `${new Date().toISOString()} | ${email} | userId:${userId || 'unknown'}\n`);
  } catch (e) { console.error('Could not write email:', e.message); }
  console.log(`Email captured: ${email}`);
  res.json({ ok: true });
});

app.post('/score-targeted', async (req, res) => {
  const { cvText, research } = req.body;
  if (!cvText) return res.status(400).json({ error: 'Missing CV text' });
  try {
    const msg = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 150,
      messages: [{ role: 'user', content: `Score this CV against the company research. Reply with ONLY: [number]/100 - [one sentence explanation under 15 words]\n\nCV:\n${cvText.substring(0, 2000)}\n\nRESEARCH:\n${(research || '').substring(0, 500)}` }]
    });
    const text = msg.content[0].text.trim();
    const match = text.match(/(\d+)\/100\s*[-–]\s*(.+)/);
    if (match) {
      res.json({ score: match[1], explain: match[2].trim() });
    } else {
      const numMatch = text.match(/(\d+)/);
      res.json({ score: numMatch ? numMatch[1] : '72', explain: text.replace(/\d+\/100\s*[-–]?\s*/,'').trim() });
    }
  } catch(e) { res.status(500).json({ error: e.message }); }
});

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

  // Admin — always Pro
  const adminEmails = ['stefan.gavra.ichb@gmail.com','mcocseraph@gmail.com'];
  if (adminEmails.includes(userEmail.toLowerCase())) {
    return res.json({ isPro: true });
  }

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

// ── 404 — MUST BE LAST ────────────────────────────────────────────────────────
app.use((req, res) => {
  res.status(404).sendFile(__dirname + '/public/landing.html');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ResumeAI running on port ${PORT}`));
