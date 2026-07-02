# RFP Automation Portal

AI-powered RFP (Request for Proposal) analysis tool that automatically extracts deliverables, evaluation criteria, compliance requirements, risks, timelines, Go/No-Go recommendations, and strategic checklist evaluations from uploaded documents.

## Features

- **Dual AI Provider** — Choose Google Gemini or Groq on upload. Gemini uses `gemini-2.5-flash` with structured output via Pydantic schema. Groq uses `openai/gpt-oss-120b` with 120B params and native JSON mode
- **Drag & Drop Upload** — Supports PDF, DOCX, and TXT (up to 16MB) with optional title and provider selection
- **Executive Summary** — Auto-generated overview of the RFP
- **8-Tab Results Interface** — Deliverables, Evaluation Criteria, Compliance, Risks, Timeline, Key Requirements, Checklist (with executive summary + status summary bar + 23-item deep analysis), Go/No-Go
- **Go / No-Go Analysis** — AI-powered pursuit recommendation with score gauge (0-100), verdict badge (Go/Caution/No-Go), and 4-8 weighted factor breakdown
- **Bid Qualification Checklist** — Unified 23-item checklist across 4 departments (5 Financial, 7 Legal, 6 Operations, 5 Technical), each with 9-field deep analysis:
  - `what_rfp_says` — What the RFP states about this item
  - `rfp_evidence` — Direct quotes, section numbers, dollar amounts from the RFP
  - `status` — Go / No-Go / Escalate / Review / Caution / Not Specified in RFP
  - `risk_level` — High / Medium / Low impact on bid success
  - `reasoning` — Detailed explanation of why this status was chosen
  - `analysis` — Competitive, financial, resource, timeline, and risk implications
  - `impact_on_bid_strategy` — How this shapes pricing, teaming, partnerships
  - `mitigation_strategy` — Actionable risk reduction plan
  - `recommendation` — Specific next steps with rationale
  - **Status Summary Bar** — Color-coded aggregate counts at the top (Go / No-Go / Review / Escalate / Caution / Not Specified)
- **Departmental Compliance** — Financial, Legal, Operations, and Technical requirements matrix
- **Risk Assessment** — Identifies risks with category and severity (High/Medium/Low)
- **Timeline & Milestones** — Key dates and deadlines from the document
- **Hybrid Tab Bar** — Desktop hover reveals ultra-thin dark-purple scrollbar; click arrow scrolls tabs smoothly; mobile hides scrollbar for native swipe
- **Analyzing Overlay** — Full-screen animated overlay with step-by-step progress cycling during document processing
- **Analysis History** — Session-based history with inline rename
- **RFP Comparison** — Side-by-side comparison of two analyses at `/compare`
- **Export Options** — Download as PDF (xhtml2pdf) or JSON
- **Search & Filter** — Real-time search across all results sections
- **Dark Mode** — Toggleable theme persisted via localStorage
- **View Transitions** — Crossfade page transitions using the View Transition API
- **Zero CDN** — All assets (CSS, JS, SVG icons, favicon) self-hosted in `static/`
- **Retry Logic** — Exponential backoff on Gemini 503 errors; no retry on auth errors or Groq 413
- **Document Truncation** — Gemini: 5000 chars; Groq: 2000 chars (fits under 8K TPM developer limit)

## Scoring Rules

- **Go/No-Go Score:** 70+ = Go, 40-69 = Caution, <40 = No-Go
- **Financial Checklists:** NET30 payment terms = Go, longer = Escalate; $5M insurance = Go, over $5M = No-Go
- **Checklist Statuses:** Go, No-Go, Escalate, Review, Caution, Not Specified in RFP

## Tech Stack

- **Backend:** Python, Flask
- **AI:** Google Gemini API (`gemini-2.5-flash`) or Groq (`openai/gpt-oss-120b` via LPU)
- **PDF Export:** xhtml2pdf
- **Document Parsing:** pypdf, python-docx
- **Models & Validation:** Pydantic (used as Gemini response_schema)
- **Frontend:** Vanilla CSS (CSS variables), Vanilla JS — no frameworks

## Setup

1. Clone the repository:
   ```bash
   git clone <repo-url>
   cd rfp-automation-portal
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install flask python-docx pypdf pydantic google-genai xhtml2pdf python-dotenv groq PyMuPDF
   ```

4. Create a `.env` file with your API keys:
   ```env
   GEMINI_API_KEY=your-gemini-api-key
   GROQ_API_KEY=your-groq-api-key
   ```
   - Get a Gemini key at [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
   - Get a Groq key at [console.groq.com/keys](https://console.groq.com/keys) (recommended for speed)

5. Run the application:
   ```bash
   .venv\Scripts\python app.py
   ```

6. Open `http://localhost:5000` in your browser.

## Usage

1. Upload an RFP (PDF, DOCX, or TXT) via drag & drop or file picker — optionally set a title and choose Gemini or Groq
2. The analyzing overlay shows progress through extraction → analysis → compliance → finalization
3. Browse results in the 8-tab interface:
   - **Deliverables** — Required tangible items and reports
   - **Evaluation** — Scoring criteria and point allocations
   - **Compliance** — Financial, Legal, Operations, Technical requirements
   - **Risks** — Categorized risks with severity ratings
   - **Timeline** — Important dates and deadlines
   - **Key Req.** — Top prioritized requirements
   - **Checklist** — 23-item deep analysis with status summary bar, RFP evidence, risk levels, impact on strategy, mitigation, and recommendations
   - **Go / No-Go** — Score gauge, verdict badge, and weighted factor reasons
4. Use the search bar to filter across all sections
5. Export as PDF or JSON
6. View analysis history at `/history` — rename analyses inline
7. Compare two analyses side by side at `/compare`
8. Toggle dark mode with the theme button (persists across sessions)

## Project Structure

```
rfp-automation-portal/
├── app.py                   # Flask routes, Pydantic schemas, AI analysis functions
├── .env                     # API keys (not committed)
├── static/
│   ├── css/theme.css        # CSS variables, base styles, page transitions
│   ├── js/theme.js          # Dark/light mode toggle
│   └── icons/
│       ├── sprite.svg       # SVG icon sprite
│       └── logo-icon.svg    # Brand favicon
├── templates/
│   ├── index.html           # Upload page with drag-drop, provider toggle, analyzing overlay
│   ├── results.html         # 8-tab results, search, export, dark mode, hybrid tab bar
│   ├── history.html         # Session history with inline rename
│   ├── compare.html         # Side-by-side document comparison
│   └── report_template.html # PDF export template with strategic checklist
├── uploads/                 # Uploaded documents (auto-created)
├── requirements.txt
├── LICENSE (MIT)
└── README.md
```
