import os
import io
import re
import json
import uuid
import hashlib
import time
import base64
from datetime import datetime
from collections import defaultdict

import pypdf
import fitz
import docx
import groq
from dotenv import load_dotenv

load_dotenv()

def _get_env(key: str) -> str:
    v = os.getenv(key, '')
    return v.strip().strip('"').strip("'") if v else v

from flask import Flask, render_template, request, make_response, Response, session
from werkzeug.utils import secure_filename
from pydantic import BaseModel, Field
from typing import List, Union
from google import genai
from google.genai import types
from google.genai.errors import ServerError, ClientError
from xhtml2pdf import pisa

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['TEMPLATES_AUTO_RELOAD'] = True
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['SESSION_TYPE'] = 'filesystem'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# In-memory history store (per session)
analysis_history = defaultdict(list)

MAX_DOC_CHARS = 5000  # Truncate to stay under Groq developer plan TPM limits

class GoNoGoReason(BaseModel):
    factor: str = Field(description="Factor or criterion being evaluated (e.g. Strategic Alignment, Resource Availability, Competitive Position)")
    detail: str = Field(description="Why this factor supports or opposes pursuing the RFP")
    weight: str = Field(description="Weight: High, Medium, or Low")

class GoNoGo(BaseModel):
    score: int = Field(description="Overall pursuit score from 0 to 100, where 0 = definite no-go and 100 = definite go")
    verdict: str = Field(description="One of: Go, Caution, No-Go")
    summary: str = Field(description="One-sentence summary of the go/no-go recommendation")
    reasons: List[GoNoGoReason] = Field(description="List of 4-8 factors that influenced the score")

class ChecklistItem(BaseModel):
    item: str = Field(description="The specific requirement or checklist item evaluated")
    status: str = Field(description="One of: Go, No-Go, Escalate, Review, Caution, Not Specified in RFP")
    reasoning: str = Field(description="Brief clear explanation for the status based on RFP content")
    what_rfp_says: str = Field(description="What the RFP document actually states or implies about this item. Quote key phrases if available. If not mentioned, state 'Not addressed in RFP.'")
    rfp_evidence: str = Field(description="Specific evidence from the RFP document: direct quotes, section numbers, clause references, page numbers, exhibit references, or specific dollar amounts and deadlines mentioned. This must be grounded in the actual document text.")
    impact_on_bid_strategy: str = Field(description="How this checklist item directly impacts our bid approach: pricing strategy, resource allocation, team composition, partnership needs, timeline implications, competitive positioning, or go/no-go decision.")
    risk_level: str = Field(description="Risk level this item poses to bid success: High (could disqualify or cause significant loss), Medium (requires careful management), Low (manageable or standard requirement).")
    mitigation_strategy: str = Field(description="Specific actionable steps to reduce or manage the risk associated with this item. Include who should handle it, what preparation is needed, and contingency plans if applicable.")
    analysis: str = Field(description="Deep analysis of implications, risks, and strategic considerations. Connect this item to overall bid qualification, competitive landscape, resource feasibility, and financial impact.")
    recommendation: str = Field(description="Specific actionable recommendation: what to do next, who to involve, what to prepare, or why to proceed or decline.")

class StrategicChecklist(BaseModel):
    executive_summary: str = Field(description="Overall recommendation summary with Go/No-Go/Escalate verdict and key reasoning")
    financial: List[ChecklistItem] = Field(description="Financial/Accounting checklist evaluation covering payment terms, financial stability, insurance, profitability, bid bond")
    legal: List[ChecklistItem] = Field(description="Legal checklist evaluation covering eligibility, capability, compliance, state registration, e-verify, contractual obligations")
    operations: List[ChecklistItem] = Field(description="Operations checklist evaluation covering required forms, deadlines, document compliance, signatory authority, vendor registration")
    technical: List[ChecklistItem] = Field(description="Technical checklist evaluation covering scope alignment, technical requirements, industry standards, security, integration needs")

class RiskItem(BaseModel):
    category: str = Field(description="Risk category (e.g. Financial, Timeline, Technical, Legal)")
    description: str = Field(description="Description of the specific risk")
    severity: str = Field(description="Severity: High, Medium, or Low")

class TimelineMilestone(BaseModel):
    milestone: str = Field(description="Milestone or deadline mentioned")
    date_reference: str = Field(description="Date or timeframe reference from the document")

class ComplianceChecklist(BaseModel):
    Financial: str = Field(description="Requirements regarding payment terms, financial stability, insurance limits, profitability, or bid bonds.")
    Legal: str = Field(description="Requirements regarding eligibility, capability, quantum of input, compliance, state registration, e-verify, or contractual obligations.")
    Operations: str = Field(description="Requirements regarding required forms, submission deadlines, document compliance, signatory authority, or vendor registration.")
    Technical: str = Field(description="Requirements regarding scope of services, technical specifications, industry standards, security, or integration.")

class RFPAnalysis(BaseModel):
    summary: str = Field(description="Executive summary of the RFP in 2-3 sentences")
    deliverables: List[str] = Field(description="List of tangible items, products, software features, or reports we must provide.")
    evaluation_criteria: List[str] = Field(description="Summary of scoring metrics, point allocations, or judgment guidelines.")
    compliance: ComplianceChecklist
    risks: List[RiskItem] = Field(description="Key risks identified in the RFP")
    timeline: List[TimelineMilestone] = Field(description="Important dates, milestones, and deadlines")
    key_requirements: List[str] = Field(description="Top 5 most important requirements from the RFP")
    go_nogo: GoNoGo = Field(description="Go/No-Go recommendation with score, verdict, summary, and supporting reasons")
    strategic_checklist: StrategicChecklist = Field(description="Strategic RFP checklist evaluation across Financial, Legal, Operations, and Technical categories")

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    with open(file_path, "rb") as f:
        reader = pypdf.PdfReader(f)
        for i, page in enumerate(reader.pages):
            if i >= 100:
                break
            content = page.extract_text()
            if content:
                text += content + "\n"
    return text

def extract_text_from_docx(file_path: str) -> str:
    doc = docx.Document(file_path)
    return "\n".join([paragraph.text for paragraph in doc.paragraphs])

def get_document_text(file_path: str) -> str:
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.endswith('.docx'):
        return extract_text_from_docx(file_path)
    else:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()

def compute_file_hash(file_path: str) -> str:
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hasher.update(chunk)
    return hasher.hexdigest()[:16]

def analyze_rfp(file_path: str) -> dict:
    gemini_key = _get_env('GEMINI_API_KEY')
    if not gemini_key:
        raise ValueError("GEMINI_API_KEY is not set in .env file.")
    document_text = get_document_text(file_path)[:MAX_DOC_CHARS]

    client = genai.Client(api_key=gemini_key)

    system_instruction = (
        "You are an elite RFP (Request for Proposal) Strategic Bid Qualification Analyst. "
        "Your role is to extract EVERY relevant detail from the RFP text and perform a deep, "
        "evidence-based analysis. Accuracy and specificity are paramount — base EVERY claim "
        "on actual text from the provided RFP document.\n\n"
        "For the go_nogo field: evaluate strategic alignment, resource fit, competitive landscape, "
        "risk exposure, and financial viability. Score 0-100 where 70+ is Go, 40-69 is Caution, "
        "below 40 is No-Go. Provide 4-8 specific evidence-backed reasons.\n\n"
        "=== STRATEGIC CHECKLIST — MANDATORY INSTRUCTIONS ===\n"
        "Evaluate the RFP against the exact checklist items below. For EVERY item, populate "
        "ALL fields with substantive, multi-sentence content. Superficial one-liners are "
        "unacceptable. Each field must demonstrate rigorous analysis:\n\n"
        "- what_rfp_says (2-5 sentences): Summarize what the RFP explicitly states about this "
        "item. Reference specific clauses, sections, requirements. If genuinely not addressed, "
        "state 'Not addressed in the RFP' and explain what a typical RFP would require.\n\n"
        "- rfp_evidence (3-7 sentences): Cite SPECIFIC evidence from the RFP — direct quotes "
        "from the document text (use quotation marks), section numbers, clause references, "
        "dollar amounts, page references, exhibit letters, submission requirements, deadlines, "
        "formatting specs, certification requirements. This MUST be grounded in the actual "
        "text. If the RFP is silent, explain what evidence is missing and why it matters.\n\n"
        "- status: One of Go, No-Go, Escalate, Review, Caution, Not Specified in RFP. "
        "In the reasoning field, EXPLAIN in detail WHY this status was chosen. Cite the "
        "specific clause or requirement that drove the decision.\n\n"
        "- risk_level: High / Medium / Low — based on impact to bid success. Justify in your "
        "analysis text.\n\n"
        "- analysis (4-8 sentences): DEEP strategic analysis — connect this item to overall "
        "bid qualification. Structure your analysis as follows:\n"
        "  (a) Competitive implication: Does this create advantage or disadvantage for us?\n"
        "  (b) Resource assessment: What people, budget, technology, or partnerships are needed?\n"
        "  (c) Financial impact: Cost to comply vs. cost of non-compliance. Profit margin impact.\n"
        "  (d) Timeline pressure: Does this constrain our bid preparation or delivery schedule?\n"
        "  (e) Hidden risks: What indirect consequences exist (e.g., disqualification, penalties, "
        "reputation damage, legal exposure)?\n"
        "  (f) Capability alignment: How well does this match our current strengths?\n\n"
        "- impact_on_bid_strategy (3-6 sentences): How this item shapes our bid approach — "
        "pricing strategy (premium vs. aggressive), team composition (internal vs. subcontractors), "
        "partnership decisions (who to team with), go/no-go tipping point, resource allocation "
        "priorities, timeline planning (early start needed?). Be specific about strategic "
        "trade-offs and decisions this item forces.\n\n"
        "- mitigation_strategy (3-6 sentences): Concrete action plan to address risks — who "
        "owns the response (role/title), what preparation is needed (documents, approvals, "
        "certifications), timeline for resolution (before submission vs. post-award), "
        "contingency plans if primary approach fails, cost of mitigation vs. cost of risk. "
        "Must be actionable, not theoretical.\n\n"
        "- recommendation (3-5 sentences): Specific actionable next steps with clear rationale. "
        "State what to do, who should do it, by when, and why it matters. Connect to overall "
        "bid strategy.\n\n"
        "CRITICAL — USE THE EXACT ITEM NAMES BELOW. Do NOT rename, split, or merge items.\n"
        "=== EXACT CHECKLIST ITEMS ===\n"
        "FINANCIAL (5 items): Payment Terms (NET30=Go, >NET30=Escalate to Accounting), "
        "Financial Stability Requirements (financial statements/proof required, unaudited acceptable?), "
        "Insurance Requirements ($5M=Go, >$5M=No-Go), Profitability Analysis (expected revenue vs projected costs), "
        "Bid Bond (required and terms).\n"
        "LEGAL (7 items): Eligibility Criteria (Relevant Experience, Registration Requirement, "
        "Financial Statement of Previous Year), Capability (Qualified Personnel, Technical Knowhow), "
        "Quantum of Input Required (Expected Revenue Generation, Period of Implementation, Insurance Coverage, "
        "Compliance of Law), Compliance Requirements (comply with relevant laws including data protection), "
        "State Registration (registered in project state), E-Verify (required?), "
        "Contractual Obligations (termination clauses, liability limits, dispute resolution).\n"
        "OPERATIONS (6 items): Required Forms (Insurance Requirement, Information Form with Tax ID/Owner "
        "Name/% ownership, Small Business MD, MBE specify type, Workers Comp Insurance, Business with Iran), "
        "Submission Deadlines (completed accurately and submitted on time), Document Compliance "
        "(formatting and submission requirements), Signatory Authority (correct individuals with authority), "
        "Checklist of Required Documents (Responsible Person: RFP Owner/Lead, Meeting with Ops), "
        "Vendor Registration (Specific Info to finish registration, who will be responsible).\n"
        "TECHNICAL (5 items): Scope of Services/Products (aligns with SPS offerings like IAM, cybersecurity), "
        "Technical Requirements (match SPS capabilities and offerings), "
        "Compliance with Industry Standards (adhere to standards and best practices), "
        "Security Considerations (data protection, encryption, access controls), "
        "Integration Needs (integration with other systems, can SPS support).\n\n"
        "STATUS RULES (with detailed why-reasoning in the reasoning field):\n"
        "- 'Go': RFP explicitly addresses the item favorably. Reason: cite the specific "
        "clause or terms that make it a Go.\n"
        "- 'No-Go': RFP explicitly imposes a requirement the bidder cannot meet (e.g. "
        "insurance > $5M, impossible deadline). Reason: state exactly what requirement "
        "cannot be satisfied and why.\n"
        "- 'Escalate': Terms exceed standard thresholds (NET>30, insurance>$5M, etc.). "
        "Reason: name the threshold and what internal approval is needed.\n"
        "- 'Review': Requirements exist but need internal assessment of capability or "
        "resources. Reason: specify what needs to be evaluated internally.\n"
        "- 'Caution': Item presents moderate risk that is manageable with attention. "
        "Reason: describe the risk and why it's manageable.\n"
        "- 'Not Specified in RFP': RFP does not address this item at all. Use this "
        "neutrally — it is NOT a No-Go by default. Reason: simply state the RFP is "
        "silent on this item.\n"
        "MANDATORY: You MUST search the RFP text THOROUGHLY before concluding 'Not Specified'. "
        "Many items have implicit evidence — payment terms may be embedded in a 'Billing' "
        "section, insurance in 'Indemnification', deadlines in 'Schedule'. Extract every "
        "relevant sentence. Only use 'Not Specified' when absolutely nothing in the text "
        "relates to the item.\n"
        "Never fabricate information. Every analysis must trace back to specific RFP text. "
        "Superficial or generic responses will be rejected."
    )

    last_exception = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model='gemini-2.0-flash',
                contents=f"Here is the RFP document text:\n\n{document_text}",
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json",
                    response_schema=RFPAnalysis,
                    temperature=0.1
                ),
            )
            response_text = response.text if response.text is not None else "{}"
            return json.loads(response_text)
        except ClientError as e:
            raise  # auth/permission errors should not be retried
        except ServerError as e:
            last_exception = e
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)
        except Exception as e:
            last_exception = e
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)

def _clean_json(text: str) -> str:
    text = text.strip()
    # Strip markdown fences
    text = re.sub(r'^```(?:json)?\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
    # Fix trailing quote before closing braces common in Groq output
    text = re.sub(r'"\s*}', '}', text)
    text = re.sub(r'"\s*]', ']', text)
    # Remove control characters (except \t, \n, \r) that break json.loads
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return text

def _extract_failed_json(exc: Exception) -> str | None:
    try:
        body = getattr(exc, 'body', {})
        if isinstance(body, str):
            body = json.loads(body)
        fg = body.get('error', {}).get('failed_generation', '')
        if fg:
            return _clean_json(fg)
    except Exception:
        pass
    return None

def analyze_rfp_groq(file_path: str) -> dict:
    api_key = _get_env('GROQ_API_KEY')
    if not api_key:
        raise ValueError("GROQ_API_KEY is not set. Please add it to your .env file.")

    document_text = get_document_text(file_path)[:1000]

    client = groq.Groq(api_key=api_key)

    system_prompt = (
        "You are an elite RFP Strategic Bid Qualification Analyst. "
        "Return ONLY valid JSON. NO markdown. NO code fences.\n"
        'Schema: {summary:string, deliverables:[string], evaluation_criteria:[string], '
        'compliance:{Financial:string, Legal:string, Operations:string, Technical:string}, '
        'risks:[{category:string, description:string, severity}], '
        'timeline:[{milestone:string, date_reference:string}], '
        'key_requirements:[string], '
        'go_nogo:{score:int(0-100), verdict, summary:string, reasons:[{factor:string, detail:string, weight}]}, '
        'strategic_checklist:{executive_summary:string, '
        'financial:[{item, status, reasoning:string, what_rfp_says:string, rfp_evidence:string, impact_on_bid_strategy:string, risk_level, mitigation_strategy:string, analysis:string, recommendation:string}], '
        'legal same-structure, operations same-structure, technical same-structure}\n\n'
        "EACH checklist item needs ALL 9 fields:\n"
        "- what_rfp_says: What the RFP states. Quote clauses/sections.\n"
        "- rfp_evidence: Specific RFP quotes with quotation marks, section numbers, dollar amounts.\n"
        "- status: Go, No-Go, Escalate, Review, Caution, or Not Specified in RFP\n"
        "- risk_level: High, Medium, or Low\n"
        "- reasoning: WHY this status. Cite the specific RFP clause.\n"
        "- analysis: Deep strategic analysis covering competitive, resource, financial, timeline, hidden risks, capability alignment.\n"
        "- impact_on_bid_strategy: How this shapes pricing, teaming, partnerships, timeline.\n"
        "- mitigation_strategy: Actionable plan with owner, timeline, contingency.\n"
        "- recommendation: What to do, who, when, why.\n\n"
        "EXACT ITEMS (do NOT rename):\n"
        "FINANCIAL (5): Payment Terms, Financial Stability Requirements, Insurance Requirements, "
        "Profitability Analysis, Bid Bond\n"
        "LEGAL (7): Eligibility Criteria, Capability, Quantum of Input Required, "
        "Compliance Requirements, State Registration, E-Verify, Contractual Obligations\n"
        "OPERATIONS (6): Required Forms, Submission Deadlines, Document Compliance, "
        "Signatory Authority, Checklist of Required Documents, Vendor Registration\n"
        "TECHNICAL (5): Scope of Services/Products, Technical Requirements, "
        "Compliance with Industry Standards, Security Considerations, Integration Needs\n\n"
        "STATUS RULES:\n"
        "- Go: RFP explicitly favorable. Reason: cite clause.\n"
        "- No-Go: RFP imposes impossible requirement. Reason: what cannot be met.\n"
        "- Escalate: Exceeds thresholds (NET>30, insurance>$5M). Reason: name threshold.\n"
        "- Review: Needs internal assessment. Reason: what needs evaluation.\n"
        "- Caution: Moderate manageable risk. Reason: describe the risk.\n"
        "- Not Specified in RFP: RFP is silent. Use only when absolutely nothing relates.\n"
        "Search thoroughly before defaulting to Not Specified. "
        "Score: 70+=Go, 40-69=Caution, <40=No-Go. Never fabricate."
    )

    last_exception = None
    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model='openai/gpt-oss-120b',
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"RFP Document Text:\n\n{document_text}"},
                ],
                response_format={"type": "json_object"},
                temperature=0.1,
                max_tokens=6000,
            )
            text = response.choices[0].message.content or "{}"
            parsed = json.loads(_clean_json(text))
            RFPAnalysis(**parsed)
            return parsed
        except (json.JSONDecodeError, Exception) as e:
            # Try extracting JSON from Groq's failed_generation
            recovered = _extract_failed_json(e)
            if recovered:
                try:
                    parsed = json.loads(recovered)
                    RFPAnalysis(**parsed)
                    return parsed
                except Exception:
                    pass
            if hasattr(e, 'status_code') and e.status_code in (400, 413):
                raise  # don't retry JSON validation or size errors
            last_exception = e
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)

@app.route('/')
def index() -> str:
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file() -> Union[Response, str]:
    if 'rfp_file' not in request.files:
        return make_response("No file part in the request", 400)

    file = request.files['rfp_file']
    if not file or file.filename == '':
        return make_response("No file selected", 400)

    filename = secure_filename(str(file.filename))
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    try:
        provider = request.form.get('provider', 'gemini')
        if provider == 'groq':
            analysis_results = analyze_rfp_groq(file_path)
        else:
            analysis_results = analyze_rfp(file_path)
        file_hash = compute_file_hash(file_path)
        word_count = len(get_document_text(file_path).split())

        analysis_title = request.form.get('analysis_title', '').strip() or filename.rsplit('.', 1)[0].replace('-', ' ').replace('_', ' ')

        record = {
            'id': str(uuid.uuid4())[:8],
            'title': analysis_title,
            'filename': filename,
            'timestamp': datetime.now().isoformat(),
            'hash': file_hash,
            'word_count': word_count,
            'results': analysis_results,
        }

        if 'session_id' not in session:
            session['session_id'] = str(uuid.uuid4())
        sid = session['session_id']
        history = analysis_history[sid]
        history.insert(0, record)
        if len(history) > 20:
            history.pop()

        return render_template('results.html', results=analysis_results, filename=filename, record_id=record['id'], record_title=analysis_title, provider=provider)
    except Exception as e:
        error_msg = str(e)
        if '503' in error_msg or 'UNAVAILABLE' in error_msg or 'ServiceUnavailable' in error_msg:
            return make_response(
                "The AI service is temporarily overloaded (503). "
                "Please wait a few seconds and try uploading again. "
                "The system will automatically retry up to 3 times on subsequent attempts.",
                503
            )
        if '413' in error_msg or 'Payload Too Large' in error_msg or 'content size limit' in error_msg.lower():
            return make_response(
                "The document is too large for the Groq API. Try a shorter document or switch to Gemini. "
                "(413 Payload Too Large)",
                413
            )
        if '429' in error_msg or 'rate_limit' in error_msg.lower() or 'Too Many Requests' in error_msg:
            return make_response(
                "API rate limit exceeded (429). Please wait 30-60 seconds then try again.",
                429
            )
        return make_response(f"An error occurred during AI analysis: {error_msg}", 500)

@app.route('/export_pdf', methods=['POST'])
def export_pdf():
    results_raw = request.form.get('results_data', '{}')
    filename = request.form.get('filename', 'document')
    try:
        results_dict = json.loads(results_raw)
    except Exception:
        results_dict = {}

    html = render_template('report_template.html', results=results_dict, filename=filename)
    buf = io.BytesIO()
    status = pisa.CreatePDF(html, dest=buf)
    if getattr(status, 'err', False):
        buf.close()
        print('PDF generation error - status.err is True')
        return make_response("Error creating PDF", 500)
    pdf_bytes = buf.getvalue()
    buf.close()
    safe_name = re.sub(r'[^\w\s-]', '', filename).strip().replace(' ', '_') or 'analysis'
    import time
    ts = str(int(time.time()))
    resp = make_response(pdf_bytes)
    resp.headers['Content-Type'] = 'application/pdf'
    resp.headers['Content-Disposition'] = f'attachment; filename="RFP_Analysis_{safe_name}_{ts}.pdf"'
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/export_json', methods=['POST'])
def export_json():
    results_raw = request.form.get('results_data', '{}')
    filename = request.form.get('filename', 'document')

    try:
        results_dict = json.loads(results_raw)
    except Exception:
        results_dict = {}

    safe_name = re.sub(r'[^\w\s-]', '', filename).strip().replace(' ', '_') or 'analysis'
    clean_name = f"RFP_Analysis_{safe_name}.json"
    response = make_response(json.dumps(results_dict, indent=2))
    response.headers['Content-Type'] = 'application/json'
    response.headers['Content-Disposition'] = f'attachment; filename="{clean_name}"'
    return response

@app.route('/history')
def history():
    sid = session.get('session_id', 'default')
    records = analysis_history.get(sid, [])
    return render_template('history.html', records=records)

@app.route('/compare', methods=['POST'])
def compare():
    record_id_1 = request.form.get('record_1')
    record_id_2 = request.form.get('record_2')
    sid = session.get('session_id', 'default')
    records = analysis_history.get(sid, [])

    r1 = next((r for r in records if r['id'] == record_id_1), None)
    r2 = next((r for r in records if r['id'] == record_id_2), None)

    if not r1 or not r2:
        return make_response("One or both records not found", 404)

    return render_template('compare.html', r1=r1, r2=r2)

@app.route('/rename', methods=['POST'])
def rename():
    record_id = request.form.get('record_id')
    new_title = request.form.get('title', '').strip()
    if not record_id or not new_title:
        return make_response("Missing parameters", 400)
    sid = session.get('session_id', 'default')
    records = analysis_history.get(sid, [])
    for r in records:
        if r['id'] == record_id:
            r['title'] = new_title
            return '', 204
    return make_response("Record not found", 404)

if __name__ == '__main__':
    app.run(debug=True)
