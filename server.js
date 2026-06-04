const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const Stripe = require('stripe');
require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const stripe = Stripe(process.env.STRIPE_SECRET_KEY);

app.use(cors());
app.use('/webhook', express.raw({ type: 'application/json' }));
app.use(express.json({ limit: '10mb' }));

// Routes must be before static middleware
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/landing.html');
});

app.get('/app', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.use(express.static('public'));

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const STUDENT_SYSTEM = `You are an expert career coach specializing in student applications — internships, spring weeks, graduate schemes, and entry-level roles at top firms (finance, consulting, tech). You understand that students have limited work experience and need to maximize extracurriculars, academics, societies, and transferable skills. You write in a confident, achievement-focused style appropriate for competitive programs.`;

const PROFESSIONAL_SYSTEM = `You are an expert executive resume writer and career strategist with 20 years of experience placing candidates at Fortune 500 companies. You specialize in quantifying achievements, positioning career pivots, and crafting narratives that pass ATS systems while impressing human recruiters. You write in a results-driven, metrics-focused style.`;

const TARGETED_SYSTEM = `You are an elite career strategist who specializes in crafting highly targeted applications for specific companies and programs. You have deep knowledge of investment banking, consulting, tech, and other industries. You use web search to find the most current and specific information about each program, then craft applications that speak directly to that company's values, culture, and selection criteria.`;

const buildPrompt = (resume, jobDescription, userType, templateStyle) => `
Analyze this resume against the job description and provide a comprehensive optimization.

USER TYPE: ${userType === 'student' ? 'Student applying for internship/spring week/grad scheme' : 'Professional applying for full-time role'}
TEMPLATE STYLE: ${templateStyle}

Respond in EXACTLY this format:

---ATS SCORE---
[number]/100 - [one sentence explanation]

---MISSING KEYWORDS---
[comma separated list of keywords from job description missing in resume]

---OPTIMIZED RESUME---
[Full optimized resume in ${templateStyle} style.]

---COVER LETTER---
[Tailored cover letter.]

---INTERVIEW TIPS---
[3 bullet points of specific interview tips based on this job description]

RESUME:
${resume}

JOB DESCRIPTION:
${jobDescription}
`;

const buildTargetedPrompt = (resume, target) => `
The user wants to apply to: "${target}"

First, use your web search tool to find:
1. Exact details about this program/role
2. The company's current values, culture, and strategic priorities
3. What makes a STRONG application for this specific program
4. Any recent news about the company relevant to the application
5. Typical interview process and what they assess

Then, using that research AND the user's resume below, create a complete application package.

Respond in EXACTLY this format:

---RESEARCH SUMMARY---
[3-4 sentences summarizing what you found about this program/company]

---APPLICATION STRATEGY---
[2-3 sentences on the specific angle to take for this application]

---TAILORED CV---
[Full CV rewritten to target this specific company and program]

---COVER LETTER---
[Cover letter specifically written for this company and program]

---INTERVIEW PREP---
[5 likely interview questions for this specific program with suggested answer frameworks]

USER'S CURRENT RESUME:
${resume}
`;

// Standard optimize
app.post('/optimize', async (req, res) => {
  const { resume, jobDescription, userType, templateStyle } = req.body;
  if (!resume || !jobDescription) return res.status(400).json({ error: 'Missing fields' });
  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4000,
      system: userType === 'student' ? STUDENT_SYSTEM : PROFESSIONAL_SYSTEM,
      messages: [{ role: 'user', content: buildPrompt(resume, jobDescription, userType, templateStyle) }]
    });
    res.json({ result: message.content[0].text });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Targeted application (Pro)
app.post('/targeted', async (req, res) => {
  const { resume, target } = req.body;
  if (!resume || !target) return res.status(400).json({ error: 'Missing fields' });
  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 6000,
      system: TARGETED_SYSTEM,
      tools: [{ type: 'web_search_20250305', name: 'web_search' }],
      messages: [{ role: 'user', content: buildTargetedPrompt(resume, target) }]
    });
    const fullText = message.content
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('\n');
    res.json({ result: fullText });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Create Stripe checkout session
app.post('/create-checkout', async (req, res) => {
  const { userId, userEmail } = req.body;
  if (!userId || !userEmail) return res.status(400).json({ error: 'Missing user info' });

  try {
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      mode: 'subscription',
      customer_email: userEmail,
      line_items: [{
        price: process.env.STRIPE_PRICE_ID,
        quantity: 1,
      }],
      metadata: { userId },
      success_url: `${process.env.APP_URL}/?upgrade=success`,
      cancel_url: `${process.env.APP_URL}/?upgrade=cancelled`,
    });
    res.json({ url: session.url });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Check subscription status
app.post('/check-subscription', async (req, res) => {
  const { userEmail } = req.body;
  if (!userEmail) return res.status(400).json({ isPro: false });

  try {
    const customers = await stripe.customers.list({ email: userEmail, limit: 1 });
    if (customers.data.length === 0) return res.json({ isPro: false });

    const customer = customers.data[0];
    const subscriptions = await stripe.subscriptions.list({
      customer: customer.id,
      status: 'active',
      limit: 1
    });

    res.json({ isPro: subscriptions.data.length > 0 });
  } catch (error) {
    res.json({ isPro: false });
  }
});

// Stripe webhook
app.post('/webhook', async (req, res) => {
  const sig = req.headers['stripe-signature'];
  let event;
  try {
    event = stripe.webhooks.constructEvent(req.body, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    return res.status(400).send(`Webhook Error: ${err.message}`);
  }
  // Handle subscription events if needed
  res.json({ received: true });
});

// Template analyzer
app.post('/analyze-template', upload.single('template'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const fileContent = fs.readFileSync(req.file.path, 'utf8');
    const message = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 500,
      messages: [{ role: 'user', content: `Analyze this resume template and describe its style in 2-3 sentences:\n\n${fileContent}` }]
    });
    fs.unlinkSync(req.file.path);
    res.json({ style: message.content[0].text });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ResumeAI running on port ${PORT}`));
