"""
Careon Services — NDIS Progress Note Generator
-----------------------------------------------
A Streamlit web app that turns rough session notes into
professional NDIS-compliant progress notes using Claude AI.

Coordinators open this in a browser — no coding required.
"""

import streamlit as st
import anthropic
import os
from datetime import date

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Progress Note Generator | Careon Services",
    page_icon="📝",
    layout="centered"
)

# ── Styling ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    h1 { color: #1a3c5e; font-size: 1.7rem !important; }
    .stTextArea textarea { font-size: 0.95rem; }
    .note-output {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1.5rem;
        line-height: 1.7;
    }
    .footer { color: #999; font-size: 0.8rem; text-align: center; margin-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.title("📝 Progress Note Generator")
st.caption("Careon Services — NDIS Support Coordination")
st.divider()

# ── API Key ──────────────────────────────────────────────────────────────────
# Tries: Streamlit secrets → environment variable → sidebar input
api_key = ""
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except:
    api_key = ""
if os.environ.get("ANTHROPIC_API_KEY"):
    api_key = os.environ["ANTHROPIC_API_KEY"]

if not api_key:
    with st.sidebar:
        st.subheader("⚙️ Setup")
        api_key = st.text_input(
            "Claude API Key",
            type="password",
            help="Enter your Anthropic API key. Get one from console.anthropic.com"
        )
        st.caption("💡 Once set up by your admin, this will be automatic.")

# ── Coordinators ─────────────────────────────────────────────────────────────
COORDINATORS = [
    "Gamal Mohamed Ali",
    "Gamil Abdalla",
    "Nemat Mohamed Ali",
]

# ── Form ─────────────────────────────────────────────────────────────────────
with st.form("note_form", clear_on_submit=False):

    col1, col2 = st.columns(2)

    with col1:
        participant_name = st.text_input(
            "Participant Name *",
            placeholder="e.g. Ahmed Aamer"
        )
        session_date = st.date_input(
            "Date of Session",
            value=date.today()
        )
        duration = st.text_input(
            "Duration",
            placeholder="e.g. 45 minutes",
            value="60 minutes"
        )

    with col2:
        coordinator = st.selectbox(
            "Support Coordinator",
            options=COORDINATORS
        )
        contact_method = st.selectbox(
            "Method of Contact",
            options=[
                "Phone Call",
                "Video Call (Microsoft Teams)",
                "In Person — Participant's Home",
                "In Person — Careon Office",
                "Email",
                "Home Visit",
            ]
        )
        goals = st.text_input(
            "NDIS Goal/s Addressed (optional)",
            placeholder="e.g. Independent Living, Community Participation"
        )

    st.markdown("#### Session Notes *")
    raw_notes = st.text_area(
        "What happened in the session?",
        height=200,
        label_visibility="collapsed",
        placeholder=(
            "Type a rough summary of what happened — don't worry about wording or structure.\n\n"
            "Example:\n"
            "Met with Ahmed for 45 mins over the phone. Discussed his upcoming OT assessment. "
            "He's been having trouble with his current housing — stairs are an issue. "
            "Called the OT team to book assessment. Also reviewed his plan goals and he's "
            "keen to join a social group. Next step: follow up OT booking next week and "
            "research social groups in Sunshine area."
        )
    )

    st.markdown("")
    submitted = st.form_submit_button(
        "✨ Generate Progress Note",
        use_container_width=True,
        type="primary"
    )

# ── Generate ──────────────────────────────────────────────────────────────────
if submitted:
    # Validation
    errors = []
    if not participant_name.strip():
        errors.append("Participant Name is required.")
    if not raw_notes.strip():
        errors.append("Session Notes are required.")
    if not api_key:
        errors.append("Claude API Key is missing. Ask your admin to set it up, or enter it in the sidebar.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Writing your progress note…"):

            # Build the prompt
            prompt = f"""You are an experienced NDIS Support Coordination specialist writing professional progress notes for a disability services organisation called Careon Services.

Your task is to turn the rough session notes below into a properly structured, professional NDIS progress note.

Session details:
- Participant: {participant_name.strip()}
- Date: {session_date.strftime('%d/%m/%Y')}
- Duration: {duration.strip() or 'Not specified'}
- Support Coordinator: {coordinator}
- Method of Contact: {contact_method}
- NDIS Goals Addressed: {goals.strip() if goals.strip() else 'As per participant NDIS plan'}

Raw session notes from coordinator:
\"\"\"{raw_notes.strip()}\"\"\"

Write a professional NDIS progress note using EXACTLY this structure and these headings:

**Date of Support:** [date in DD/MM/YYYY format]
**Participant:** [full name]
**Support Coordinator:** [coordinator name]
**Duration:** [duration]
**Method of Contact:** [method]

**NDIS Goal/s Addressed:**
[State the relevant NDIS goals or support areas this session relates to. Be specific.]

**Support Provided:**
[2–4 sentences. Describe what the support coordinator did: discussions held, calls made, referrals initiated, coordination activities, information provided. Be professional and specific. Use past tense.]

**Participant Response:**
[1–3 sentences. How did the participant engage? Their capacity, willingness, any strengths or barriers observed. Use person-first language.]

**Outcomes Achieved:**
[1–3 sentences. What was accomplished in this session? What progress was made toward goals?]

**Next Steps / Actions Required:**
[List 2–5 specific follow-up actions with timeframes where possible. Use bullet points starting with -]

Rules:
- Use professional, NDIS-appropriate language
- Be specific and factual — do not invent information not in the raw notes
- If something is unclear, describe it in general terms rather than guessing specifics
- Use person-first language (e.g. "the participant" not "the disabled person")
- Keep it concise but comprehensive — no padding or filler sentences
- Do not add any commentary before or after the note
- Output the note only, nothing else"""

            try:
                client = anthropic.Anthropic(api_key=api_key)
                message = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=1200,
                    messages=[{"role": "user", "content": prompt}]
                )
                note_text = message.content[0].text.strip()

                st.success("✅ Progress note generated!")
                st.divider()

                # Formatted preview
                st.subheader("Generated Progress Note")
                st.markdown(note_text)
                st.divider()

                # Plain text for copying into ShiftCare
                st.subheader("📋 Copy into ShiftCare")
                st.caption("Select all the text below and paste it directly into ShiftCare.")
                plain_text = note_text.replace("**", "")
                st.text_area(
                    "Plain text",
                    value=plain_text,
                    height=380,
                    label_visibility="collapsed"
                )

                st.info(
                    "⚠️ **Before saving:** Review the note to make sure all details are accurate. "
                    "AI can make mistakes — you are responsible for the final content."
                )

            except anthropic.AuthenticationError:
                st.error("❌ Invalid API key. Please check your Claude API key and try again.")
            except anthropic.RateLimitError:
                st.error("❌ Rate limit reached. Please wait a moment and try again.")
            except Exception as e:
                st.error(f"❌ Something went wrong: {str(e)}")
                st.caption("If this keeps happening, contact your admin.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    '<div class="footer">Careon Services © 2026 &nbsp;|&nbsp; '
    'Progress notes must be reviewed for accuracy before saving to ShiftCare</div>',
    unsafe_allow_html=True
)
