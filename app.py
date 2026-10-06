import html
import streamlit as st
import re
import hashlib
import tempfile
import os
from pathlib import Path

from pypdf import PdfReader
from docx import Document
from sentence_transformers import SentenceTransformer
import chromadb


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="AI Document Intelligence Assistant",
    page_icon=str(Path(__file__).parent / "assets" / "document-assistant-logo.svg"),
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CUSTOM CSS
# ============================================================

def safe_html(value):
    text = "" if value is None else str(value)
    return html.escape(text, quote=False).replace("\n", "<br>")


st.markdown("""
<style>

    .stApp {
        background: radial-gradient(circle at top left, #eef4ff 0%, #f5f7fb 38%, #f8fafc 100%);
        color: #0f172a;
    }

    [data-testid="stHeader"] {
        background: transparent;
        border: 0;
        box-shadow: none;
    }

    [data-testid="stDecoration"] {
        display: none;
    }

    .block-container {
        padding-top: 0.75rem;
        padding-bottom: 3rem;
        max-width: 100%;
    }

    [data-testid="stAppViewContainer"] h1 {
        font-size: 26px;
        margin-bottom: 0.35rem;
        font-weight: 800;
        letter-spacing: -0.02em;
    }

    [data-testid="stCaptionContainer"] {
        font-size: 16px;
        line-height: 1.7;
        color: #475569;
    }

    .step-card {
        background: rgba(255,255,255,0.92);
        padding: 22px 20px 18px;
        border-radius: 18px;
        border: 1px solid #e2e8f0;
        min-height: 150px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.05);
    }

    .step-number {
        font-size: 28px;
        font-weight: 800;
        color: #2563eb;
    }

    .step-title {
        font-size: 18px;
        font-weight: 700;
        margin-top: 8px;
        color: #0f172a;
    }

    .step-text {
        color: #475569;
        font-size: 14px;
        margin-top: 8px;
        line-height: 1.6;
    }

    .section-title {
        font-size: 25px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 24px;
        margin-bottom: 16px;
        letter-spacing: -0.02em;
    }

    .requirement-card {
        background: white;
        border-radius: 16px;
        padding: 18px 18px 16px;
        border: 1px solid #e2e8f0;
        margin-bottom: 12px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
    }

    .requirement-name {
        font-size: 17px;
        font-weight: 700;
        color: #0f172a;
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
    }

    .evidence {
        font-size: 13px;
        color: #475569;
        margin-top: 8px;
        line-height: 1.6;
    }

    .required-badge {
        background: #fee2e2;
        color: #b91c1c;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.04em;
    }

    .conditional-badge {
        background: #fef3c7;
        color: #92400e;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.04em;
    }

    .optional-badge {
        background: #dbeafe;
        color: #1d4ed8;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.04em;
    }

    .match-card {
        background: white;
        padding: 18px 18px 16px;
        border-radius: 16px;
        border: 1px solid #e2e8f0;
        margin-bottom: 12px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
    }

    .matched {
        border-left: 5px solid #16a34a;
    }

    .missing {
        border-left: 5px solid #dc2626;
    }

    .review {
        border-left: 5px solid #f59e0b;
    }

    .duplicate {
        border-left: 5px solid #7c3aed;
    }

    .info-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        padding: 15px;
        border-radius: 12px;
        color: #1e3a8a;
        margin: 15px 0;
    }

    .warning-box {
        background: #fffbeb;
        border: 1px solid #fde68a;
        padding: 15px;
        border-radius: 12px;
        color: #92400e;
        margin: 15px 0;
    }

    .footer {
        text-align: center;
        color: #64748b;
        font-size: 13px;
        margin-top: 40px;
        padding: 20px;
    }

    [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #93c5fd;
        border-radius: 16px;
        background: rgba(239, 246, 255, 0.75);
    }

    .st-key-application_upload [data-testid="stFileUploader"] button {
        display: inline-flex;
        align-items: center;
        gap: 0.45rem;
        background: #000000 !important;
        color: #ffffff !important;
        border: 1px solid #000000 !important;
        font-size: 0 !important;
        transition: background-color 0.15s ease, border-color 0.15s ease;
    }

    .st-key-application_upload [data-testid="stFileUploader"] button > * {
        display: none !important;
    }

    .st-key-application_upload [data-testid="stFileUploader"] button::before {
        content: "⬆";
        color: #ffffff;
        font-size: 15px;
        line-height: 1;
    }

    .st-key-application_upload [data-testid="stFileUploader"] button::after {
        content: "Upload";
        color: #ffffff;
        font-size: 14px;
        font-weight: 600;
        line-height: 1;
    }

    .st-key-application_upload [data-testid="stFileUploader"] button:hover,
    .st-key-application_upload [data-testid="stFileUploader"] button:focus,
    .st-key-application_upload [data-testid="stFileUploader"] button:active {
        background: #000000 !important;
        color: #ffffff !important;
        border-color: #000000 !important;
    }

    .st-key-supporting_upload [data-testid="stFileUploaderDropzone"] {
        background: #000000 !important;
        border-color: #000000 !important;
        color: #ffffff !important;
    }

    .st-key-supporting_upload [data-testid="stFileUploaderDropzone"] *,
    .st-key-supporting_upload [data-testid="stFileUploaderFile"] * {
        color: #ffffff !important;
    }

    .st-key-supporting_upload [data-testid="stFileUploaderDropzone"] button,
    .st-key-supporting_upload [data-testid="stFileUploaderFile"] {
        background: #000000 !important;
        border-color: #475569 !important;
    }

    .st-key-extracted-application-text [data-testid="stExpander"],
    .st-key-extracted-application-text [data-testid="stExpander"] details,
    .st-key-extracted-application-text [data-testid="stExpander"] summary,
    .st-key-extracted-application-text [data-testid="stTextArea"] textarea {
        background: #000000 !important;
        color: #ffffff !important;
        border-color: #334155 !important;
    }

    .st-key-extracted-application-text [data-testid="stExpander"] *,
    .st-key-extracted-application-text [data-testid="stTextArea"] label {
        color: #ffffff !important;
    }

    div[data-testid="stMetric"] {
        border-radius: 16px;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
        border: 1px solid #e2e8f0;
    }

</style>
""", unsafe_allow_html=True)


# ============================================================
# HERO
# ============================================================

logo_column, title_column = st.columns([1, 10], vertical_alignment="center")
with logo_column:
    st.image(
        str(Path(__file__).parent / "assets" / "document-assistant-logo.svg"),
        width=88,
    )
with title_column:
    st.title("AI Document Intelligence Assistant")
st.caption("Upload any application")


# ============================================================
# WORKFLOW
# ============================================================

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">01</div>
        <div class="step-title">Upload Application</div>
        <div class="step-text">
            Upload a scholarship form, college application,
            government form, registration form or instructions PDF.
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">02</div>
        <div class="step-title">AI Extracts Requirements</div>
        <div class="step-text">
            AI reads the uploaded application and identifies
            required, conditional and optional documents.
        </div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">03</div>
        <div class="step-title">Verify Documents</div>
        <div class="step-text">
            Upload your supporting documents and check which
            requirements are matched, missing or need review.
        </div>
    </div>
    """, unsafe_allow_html=True)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("🤖 About")

    st.write(
        "AI Document Intelligence Assistant helps users understand "
        "what documents are required for an application without "
        "manually reading long forms."
    )

    st.divider()

    st.subheader("Supported Files")

    st.write("📄 PDF")
    st.write("📝 DOCX")
    st.write("🖼️ PNG")
    st.write("🖼️ JPG / JPEG")

    st.divider()

    st.subheader("Important")

    st.caption(
        "The uploaded application is treated as the primary "
        "source of requirements. The system does not use a "
        "fixed application checklist."
    )

    st.caption(
        "Document matching checks document type/content similarity. "
        "It does not verify whether a document is legally authentic."
    )


# ============================================================
# SESSION STATE
# ============================================================

if "application_text" not in st.session_state:
    st.session_state.application_text = ""

if "requirements" not in st.session_state:
    st.session_state.requirements = []

if "application_name" not in st.session_state:
    st.session_state.application_name = ""

if "supporting_files" not in st.session_state:
    st.session_state.supporting_files = []


# ============================================================
# TEXT EXTRACTION
# ============================================================

def extract_pdf_text(uploaded_file):

    text = ""

    try:
        reader = PdfReader(uploaded_file)

        for page in reader.pages:
            page_text = page.extract_text()

            if page_text:
                text += page_text + "\n"

    except Exception as e:
        st.error(f"PDF extraction error: {e}")

    return text


def extract_docx_text(uploaded_file):

    text = ""

    try:

        document = Document(uploaded_file)

        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                text += paragraph.text + "\n"

        for table in document.tables:

            for row in table.rows:

                row_text = []

                for cell in row.cells:
                    row_text.append(cell.text.strip())

                text += " | ".join(row_text) + "\n"

    except Exception as e:
        st.error(f"DOCX extraction error: {e}")

    return text


def extract_image_text(uploaded_file):

    try:

        import pytesseract
        from PIL import Image

        image = Image.open(uploaded_file)

        return pytesseract.image_to_string(image)

    except Exception as e:

        st.warning(
            "OCR could not be performed. Make sure Tesseract OCR "
            "is installed and available in PATH."
        )

        return ""


def extract_text(uploaded_file):

    extension = Path(uploaded_file.name).suffix.lower()

    if extension == ".pdf":
        return extract_pdf_text(uploaded_file)

    if extension == ".docx":
        return extract_docx_text(uploaded_file)

    if extension in [".png", ".jpg", ".jpeg"]:
        return extract_image_text(uploaded_file)

    return ""


# ============================================================
# NORMALIZATION
# ============================================================

ALIASES = {
    "aadhaar card": "Aadhaar Card",
    "aadhar card": "Aadhaar Card",
    "aadhaar": "Aadhaar Card",
    "aadhar": "Aadhaar Card",

    "pan card": "PAN Card",
    "pan": "PAN Card",

    "income certificate": "Income Certificate",
    "income proof": "Income Certificate",

    "bank passbook": "Bank Passbook",
    "bank statement": "Bank Statement",

    "fee receipt": "Fee Receipt",
    "fee payment receipt": "Fee Receipt",

    "college id": "College ID Card",
    "college identity card": "College ID Card",
    "student id": "College ID Card",

    "marks memo": "Marks Memo",
    "marksheet": "Marksheet",
    "mark sheet": "Marksheet",

    "bonafide certificate": "Bonafide Certificate",

    "caste certificate": "Caste Certificate",

    "ews certificate": "EWS Certificate",

    "domicile certificate": "Domicile Certificate",

    "residence certificate": "Residence Certificate",

    "birth certificate": "Birth Certificate",

    "passport": "Passport",

    "photograph": "Passport Photograph",
    "photo": "Passport Photograph",

    "signature": "Signature",

    "ration card": "Ration Card",

    "disability certificate": "Disability Certificate",

    "bank account details": "Bank Account Details",
}


def normalize_requirement(name):

    cleaned = re.sub(r"\s+", " ", name.strip())

    key = cleaned.lower()

    if key in ALIASES:
        return ALIASES[key]

    return cleaned


# ============================================================
# REQUIREMENT EXTRACTION
# ============================================================

REQUIRED_CUES = [
    "must",
    "shall",
    "required",
    "mandatory",
    "compulsory",
    "submit",
    "upload",
    "attach",
    "enclose",
    "provide"
]

CONDITIONAL_CUES = [
    "if",
    "only if",
    "where applicable",
    "applicable",
    "depending on",
    "applicants who",
    "in case of",
    "for candidates with"
]

OPTIONAL_CUES = [
    "optional",
    "may submit",
    "if available",
    "where available",
    "voluntary"
]

NEGATIVE_CUES = [
    "not required",
    "not needed",
    "do not submit",
    "not necessary",
    "no need to submit"
]

AMBIGUOUS_CUES = [
    "for example",
    "e.g.",
    "such as",
    "or equivalent",
    "or",
    "either",
    "one of"
]


def contains_any(text, words):

    text_lower = text.lower()

    return any(
        re.search(
            r"(?<!\w)" + re.escape(word) + r"(?!\w)",
            text_lower
        )
        for word in words
    )


def clean_sentence(sentence):

    sentence = re.sub(r"\s+", " ", sentence).strip()

    return sentence


def extract_document_mentions(sentence):

    found = []
    known_spans = []

    sentence_lower = sentence.lower()

    # First check known aliases
    for alias, canonical in ALIASES.items():

        for match in re.finditer(
            r"\b" + re.escape(alias) + r"\b",
            sentence_lower
        ):
            if canonical not in found:
                found.append(canonical)
            known_spans.append(match.span())

    # Generic document patterns
    generic_patterns = [
        r"\b[A-Za-z][A-Za-z0-9/&,\- ]{2,60}"
        r"(?:certificate|card|receipt|document|proof|"
        r"statement|passbook|letter|form|photograph|"
        r"identity|id)\b"
    ]

    for pattern in generic_patterns:

        matches = re.findall(pattern, sentence, flags=re.I)

        for match in matches:

            cleaned = normalize_requirement(match)
            cleaned = re.sub(
                r"^.*\b(separate|additional)\s+form$",
                lambda form_match: f"{form_match.group(1).title()} Form",
                cleaned,
                flags=re.IGNORECASE
            )
            cleaned = re.sub(
                r"^(?:a|an|the)\s+",
                "",
                cleaned,
                flags=re.IGNORECASE
            )

            if len(cleaned) > 3 and cleaned not in found:
                match_span = re.search(
                    re.escape(match),
                    sentence,
                    flags=re.IGNORECASE
                )
                overlaps_known_alias = (
                    match_span is not None
                    and any(
                        match_span.start() < known_end
                        and match_span.end() > known_start
                        for known_start, known_end in known_spans
                    )
                )
                if not overlaps_known_alias:
                    found.append(cleaned)

    return found


def classify_requirement(sentence):

    sentence_lower = sentence.lower()

    if contains_any(sentence_lower, NEGATIVE_CUES):
        return "excluded"

    if contains_any(sentence_lower, AMBIGUOUS_CUES):
        return "review"

    if contains_any(sentence_lower, CONDITIONAL_CUES):
        return "conditional"

    if contains_any(sentence_lower, OPTIONAL_CUES):
        return "optional"

    if contains_any(sentence_lower, REQUIRED_CUES):
        return "required"

    return None


def extract_requirements(text):

    requirements = []

    # Keep text after colons with its label so requirement cues apply to lists.
    sentences = re.split(
        r"(?<=[.!?])\s+|\n+",
        text
    )

    for sentence in sentences:

        sentence = clean_sentence(sentence)

        if len(sentence) < 10:
            continue

        category = classify_requirement(sentence)

        if category is None:
            continue

        documents = extract_document_mentions(sentence)

        for document in documents:

            if category == "excluded":
                continue

            existing = None

            for req in requirements:

                if req["name"].lower() == document.lower():
                    existing = req
                    break

            if existing:

                # Upgrade required if strong evidence appears
                if category == "required":
                    existing["category"] = "required"

                continue

            requirements.append({
                "name": document,
                "category": category,
                "evidence": sentence
            })

    return requirements


# ============================================================
# APPLICATION INFORMATION
# ============================================================

def extract_application_name(text):

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    keywords = [
        "application",
        "scholarship",
        "registration",
        "admission",
        "certificate",
        "scheme",
        "form"
    ]

    for line in lines[:30]:

        if any(word in line.lower() for word in keywords):

            if len(line) < 150:
                return line

    return "Uploaded Application"


def extract_deadlines(text):

    patterns = [

        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

        r"\b\d{1,2}\s+"
        r"(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December)"
        r"\s+\d{4}\b"
    ]

    dates = []

    for pattern in patterns:

        matches = re.findall(pattern, text, flags=re.I)

        for match in matches:

            if match not in dates:
                dates.append(match)

    return dates


# ============================================================
# EMBEDDING MODEL
# ============================================================

@st.cache_resource
def load_model():

    return SentenceTransformer("all-MiniLM-L6-v2")


model = load_model()


# ============================================================
# CHROMA DATABASE
# ============================================================

@st.cache_resource
def get_chroma_client():

    return chromadb.Client()


chroma_client = get_chroma_client()


# ============================================================
# REQUIREMENT MATCHING
# ============================================================

def create_requirement_collection(requirements):

    fingerprint = hashlib.md5(
        "|".join(
            sorted(
                req["name"].lower()
                for req in requirements
            )
        ).encode()
    ).hexdigest()

    collection_name = "req_" + fingerprint[:20]

    try:
        collection = chroma_client.get_collection(
            collection_name
        )

    except Exception:

        collection = chroma_client.create_collection(
            name=collection_name,
            metadata={"description": "Application requirements"}
        )

        names = [
            req["name"]
            for req in requirements
        ]

        embeddings = model.encode(
            names,
            normalize_embeddings=True
        ).tolist()

        collection.add(
            ids=[
                str(i)
                for i in range(len(names))
            ],
            documents=names,
            embeddings=embeddings
        )

    return collection


def match_supporting_documents(
    supporting_documents,
    requirements
):

    if not requirements:
        return []

    collection = create_requirement_collection(
        requirements
    )

    results = []

    for document in supporting_documents:

        document_text = document["text"]

        if not document_text.strip():
            document_text = document["name"]

        embedding = model.encode(
            document_text[:5000],
            normalize_embeddings=True
        ).tolist()

        query_result = collection.query(
            query_embeddings=[embedding],
            n_results=min(3, len(requirements))
        )

        if not query_result["ids"]:
            continue

        matched_ids = query_result["ids"][0]
        distances = query_result["distances"][0]

        best_id = matched_ids[0]
        best_distance = distances[0]

        matched_requirement = requirements[
            int(best_id)
        ]

        # Cosine distance -> approximate similarity
        similarity = max(
            0,
            min(
                1,
                1 - best_distance
            )
        )

        if similarity >= 0.65:
            status = "Matched"
            strength = "High"

        elif similarity >= 0.48:
            status = "Needs Review"
            strength = "Medium"

        else:
            status = "Needs Review"
            strength = "Low"

        results.append({
            "document": document["name"],
            "requirement": matched_requirement["name"],
            "similarity": similarity,
            "status": status,
            "strength": strength
        })

    return results


# ============================================================
# APPLICATION UPLOAD
# ============================================================

st.markdown(
    '<div class="section-title">📤 1. Upload Application / Form</div>',
    unsafe_allow_html=True
)

application_file = st.file_uploader(
    "Upload the application document",
    type=[
        "pdf",
        "docx",
        "png",
        "jpg",
        "jpeg"
    ],
    key="application_upload"
)


if application_file:

    with st.spinner("Reading your application..."):

        extracted_text = extract_text(
            application_file
        )

    if extracted_text.strip():

        st.session_state.application_text = extracted_text

        st.session_state.application_name = (
            extract_application_name(
                extracted_text
            )
        )

        st.session_state.requirements = (
            extract_requirements(
                extracted_text
            )
        )

        st.success(
            "Application successfully analyzed."
        )

    else:

        st.error(
            "Could not extract readable text from this file."
        )


# ============================================================
# DISPLAY APPLICATION INFORMATION
# ============================================================

if st.session_state.application_text:

    st.markdown(
        '<div class="section-title">📋 Application Analysis</div>',
        unsafe_allow_html=True
    )

    st.info(
        f"Detected application: "
        f"**{st.session_state.application_name}**"
    )

    deadlines = extract_deadlines(
        st.session_state.application_text
    )

    if deadlines:

        st.write(
            "📅 **Detected deadline/date:** "
            + ", ".join(deadlines)
        )

    requirements = st.session_state.requirements

# ============================================================
# REQUIREMENT CHECKLIST
# ============================================================

if st.session_state.requirements:

    st.markdown(
        '<div class="section-title">🧾 Documents Required for This Application</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "These requirements were extracted from the uploaded application."
    )

    for requirement in st.session_state.requirements:
        category = requirement["category"]
        category_label = {
            "required": "Required",
            "conditional": "Conditional",
            "optional": "Optional",
            "review": "Needs review",
        }.get(category, "Needs review")

        with st.container(border=True):
            st.markdown(f"**📄 {requirement['name']}**")
            st.caption(f"Category: {category_label}")
            st.text(f"Evidence: {requirement['evidence']}")


elif st.session_state.application_text:

    st.warning(
        "No clear document requirements were detected. "
        "The application may use tables, scanned images, "
        "or wording that requires manual review."
    )


# ============================================================
# SUPPORTING DOCUMENT UPLOAD
# ============================================================

if st.session_state.application_text:

    st.markdown(
        '<div class="section-title">📎 2. Upload Your Supporting Documents</div>',
        unsafe_allow_html=True
    )

    supporting_files = st.file_uploader(
        "Upload the documents you already have",
        type=[
            "pdf",
            "docx",
            "png",
            "jpg",
            "jpeg"
        ],
        accept_multiple_files=True,
        key="supporting_upload"
    )

    if supporting_files:

        documents = []

        for uploaded_file in supporting_files:

            with st.spinner(
                f"Reading {uploaded_file.name}..."
            ):

                text = extract_text(
                    uploaded_file
                )

            documents.append({
                "name": uploaded_file.name,
                "text": text
            })

        st.session_state.supporting_files = documents

        st.success(
            f"{len(documents)} supporting document(s) uploaded."
        )


# ============================================================
# VERIFICATION
# ============================================================

if (
    st.session_state.requirements
    and st.session_state.supporting_files
):

    st.markdown(
        '<div class="section-title">🔍 3. Document Verification</div>',
        unsafe_allow_html=True
    )

    verification_results = match_supporting_documents(
        st.session_state.supporting_files,
        st.session_state.requirements
    )

    # --------------------------------------------------------
    # Duplicate detection
    # --------------------------------------------------------

    hashes = {}

    duplicates = set()

    for document in st.session_state.supporting_files:

        content_hash = hashlib.md5(
            (
                document["text"]
                + document["name"]
            ).encode(
                errors="ignore"
            )
        ).hexdigest()

        if content_hash in hashes:

            duplicates.add(
                document["name"]
            )

        else:

            hashes[content_hash] = document["name"]


    # --------------------------------------------------------
    # Determine matched requirements
    # --------------------------------------------------------

    matched_requirements = set()

    review_requirements = set()

    matched_documents = []

    for result in verification_results:

        if result["status"] == "Matched":

            matched_requirements.add(
                result["requirement"]
            )

            matched_documents.append(
                result["document"]
            )

        else:

            review_requirements.add(
                result["requirement"]
            )


    required_requirements = {
        r["name"]
        for r in st.session_state.requirements
        if r["category"] == "required"
    }


    missing_requirements = (
        required_requirements
        - matched_requirements
    )


    # --------------------------------------------------------
    # Metrics
    # --------------------------------------------------------

    total_required = len(
        required_requirements
    )

    total_matched = len(
        matched_requirements
        & required_requirements
    )

    total_missing = len(
        missing_requirements
    )

    completion = (
        total_matched / total_required * 100
        if total_required
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Required",
        total_required
    )

    col2.metric(
        "Matched",
        total_matched
    )

    col3.metric(
        "Missing",
        total_missing
    )

    col4.metric(
        "Completion",
        f"{completion:.0f}%"
    )


    # --------------------------------------------------------
    # Verification results
    # --------------------------------------------------------

    st.subheader("Document Matching")

    for result in verification_results:

        if result["status"] == "Matched":

            card_class = "matched"
            icon = "✅"

        else:

            card_class = "review"
            icon = "⚠️"

        st.markdown(
            f"""
            <div class="match-card {card_class}">

                <b>{icon} {safe_html(result["document"])}</b>

                <br><br>

                <b>Matched requirement:</b>
                {safe_html(result["requirement"])}

                <br>

                <b>Match strength:</b>
                {safe_html(result["strength"])}

            </div>
            """,
            unsafe_allow_html=True
        )


    # --------------------------------------------------------
    # Missing documents
    # --------------------------------------------------------

    st.subheader("❌ Missing Required Documents")

    if missing_requirements:

        for missing in sorted(
            missing_requirements
        ):

            st.markdown(
                f"""
                <div class="match-card missing">

                    <b>❌ {safe_html(missing)}</b>

                    <br>

                    <span>
                        This required document was not
                        confidently matched with your uploads.
                    </span>

                </div>
                """,
                unsafe_allow_html=True
            )

    else:

        st.success(
            "🎉 All clearly required documents have "
            "a confident match."
        )


    # --------------------------------------------------------
    # Conditional documents
    # --------------------------------------------------------

    conditional_requirements = [
        r["name"]
        for r in st.session_state.requirements
        if r["category"] == "conditional"
    ]

    if conditional_requirements:

        st.subheader(
            "🟡 Conditional Documents"
        )

        st.caption(
            "These documents are required only when "
            "the condition mentioned in the application applies."
        )

        for item in conditional_requirements:

            st.write(
                f"• {item}"
            )


    # --------------------------------------------------------
    # Duplicate documents
    # --------------------------------------------------------

    if duplicates:

        st.subheader(
            "♻️ Possible Duplicate Files"
        )

        for duplicate in duplicates:

            st.warning(
                f"Possible duplicate: {duplicate}"
            )


# ============================================================
# FINAL CHECKLIST
# ============================================================

if st.session_state.requirements:

    st.markdown(
        '<div class="section-title">✅ Final Checklist</div>',
        unsafe_allow_html=True
    )

    for requirement in st.session_state.requirements:

        if requirement["category"] == "required":

            st.checkbox(
                requirement["name"],
                key="check_" + hashlib.md5(
                    requirement["name"].encode()
                ).hexdigest()
            )

        elif requirement["category"] == "conditional":

            st.checkbox(
                "Conditional: "
                + requirement["name"],
                key="conditional_" + hashlib.md5(
                    requirement["name"].encode()
                ).hexdigest()
            )

        else:

            st.checkbox(
                "Optional: "
                + requirement["name"],
                key="optional_" + hashlib.md5(
                    requirement["name"].encode()
                ).hexdigest()
            )


# ============================================================
# RAW APPLICATION TEXT
# ============================================================

if st.session_state.application_text:

    with st.container(key="extracted-application-text"):
        with st.expander("🔎 View Extracted Application Text"):

            st.text_area(
                "Extracted text",
                st.session_state.application_text,
                height=350
            )


# ============================================================
# DISCLAIMER
# ============================================================

st.markdown("""
<div class="warning-box">

<b>⚠️ Important:</b><br>

This application identifies document requirements and
matches uploaded files using text extraction and semantic
similarity. It does not verify the legal authenticity,
validity, ownership, or genuineness of any government,
bank, college, or identity document.

If the application contains unclear instructions,
examples, alternatives, or conditions that cannot be
reliably interpreted, manually review the original form.

</div>
""", unsafe_allow_html=True)


# ============================================================
# FOOTER
# ============================================================

st.markdown("""
<div class="footer">
    AI Document Intelligence Assistant •
    Intelligent requirement extraction and document matching
</div>
""", unsafe_allow_html=True)
