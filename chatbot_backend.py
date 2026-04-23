import os
import re
import uuid
from logger_config import get_logger
from datetime import datetime
from dotenv import load_dotenv
from masterprompt import PITCH_DECK_FIELDS, MASTER_PROMPT, WORKFLOW_QUESTIONS, AVAILABLE_CAPABILITIES, CAPABILITY_OPTIONS_STR
from masterprompt import is_no_value_answer_semantic_prompt, is_smalltalk_prompt, is_out_of_scope_prompt, is_custom_slide_request_prompt 
from masterprompt import is_restart_prompt, is_cancel_deck_prompt, is_pitch_deck_request_prompt , extract_CLEAN_RAG_PROMPT
from masterprompt import (
    _chat_once_is_out_of_scope_prompt, _chat_once_is_smalltalk_prompt, _ask_slides_gathering_permission_reply_prompt, _deck_generation_status_decline_prompt,
    _extract_custom_slide_request_prompt, _ask_llm_for_generate_control_prompt, _missing_topic_response_prompt,
    llm_intent_classify_prompt, _extract_details_from_message_prompt,
    _skip_instruction_prompt, _get_next_question_from_llm_prompt,
    _compress_history_prompt, _generate_situational_response_prompt,
    is_confirmation_prompt, SUMMARY_CAPABILITY_LABEL, SUMMARY_SLIDE_COUNT_LABEL, SUMMARY_SLIDE_COUNT_QUESTION
)
import json
import threading
from typing import Optional, Tuple, Dict, Any, List

CONTROL_JSON_REGEX = re.compile(
    r"```(?:json)?\s*(\{[\s\S]*?\})\s*```",
    flags=re.S
)


# -----------------------------
# Conversation loop guardrails
# -----------------------------
# These helpers prevent the workflow from repeatedly re-asking the same
# field question when the user is effectively saying "I don't have an answer".

_NEGATIVE_SKIP_TOKENS = {
    "no",
    "no.",
    "nope",
    "nah",
    "skip",
    "n/a",
    "none",
    "nothing",
    "not sure",
    "unsure",
    "idk",
    "i dont know",
    "i don't know",
    "do not know",
    "not applicable",
    "leave it blank",
}

_AFFIRMATION_ONLY_TOKENS = {
    "yes",
    "yep",
    "yeah",
    "yup",
    "sure",
    "ok",
    "okay",
    "absolutely",
}


def _normalize_user_text_for_intent(message: Optional[str]) -> str:
    """Normalizes a user message for simple token matching."""
    t = (message or "").strip().lower()
    # Normalize common trailing punctuation.
    if t.endswith((".", "!", "?")):
        t = t[:-1].strip()
    return t


def _is_negative_skip_message(message: Optional[str]) -> bool:
    """True when the user response is effectively 'no answer / skip'."""
    t = _normalize_user_text_for_intent(message)
    return t in _NEGATIVE_SKIP_TOKENS


def _is_affirmation_only_message(message: Optional[str]) -> bool:
    """True when the user response is essentially an affirmation (no extra info)."""
    t = _normalize_user_text_for_intent(message)
    return t in _AFFIRMATION_ONLY_TOKENS


def _get_missing_required_field(state: dict) -> Optional[str]:
    """
    Smart-gate required fields in this workflow.
    We treat fields as 'present' if they are set to '' (explicit skip), not only if non-empty.
    """
    for f in ("capability", "client", "challenge"):
        if state.get(f) is None:
            return f
    return None


# LangChain / Azure clients (expected to be available in your environment)
from langchain_openai import AzureChatOpenAI, AzureOpenAIEmbeddings
from langchain_community.vectorstores.azuresearch import AzureSearch
from langchain.prompts import PromptTemplate
from langchain.chains import ConversationalRetrievalChain

# Local agent tools (must exist in your repo)
import powerpoint_agent_tools as agent_tools
from powerpoint_agent_tools_stable import (
    get_access_token,
    get_site_and_drive_id,
)



# Initialize logger
# logging.basicConfig(
#     level=logging.INFO,
#     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
# )
# logger = logging.getLogger(__name__)
logger = get_logger(__name__)

# Load environment
load_dotenv()

# -----------------------------
# 🟢 UPDATED: Context Window Limit
# -----------------------------
MAX_HISTORY_TURNS  = 10   # Total turns kept before archiving
MAX_CONTEXT_TURNS  = 6    # Turns injected into any single LLM prompt
MAX_CONTEXT_TOKENS = 3000 # Soft token ceiling for context injection


def _estimate_tokens(text: str) -> int:
    """Rough token estimate: ~4 characters per token (safe undercount)."""
    return max(1, len(text) // 4)


def _build_context_block(history: list, max_turns: int = MAX_CONTEXT_TURNS, compressed_context: str = "") -> str:
    """
    Builds a clean, token-guarded conversation context string from recent turns.
    Iterates backwards so the most recent turns are always included first,
    stopping when the token budget (MAX_CONTEXT_TOKENS) is reached.
    """
    recent = history[-max_turns:] if history else []
    lines = []
    token_count = 0
    for turn in reversed(recent):
        u = (turn.get("user") or "").strip()
        b = (turn.get("bot")  or "").strip()
        block = f"User: {u}\nGia: {b}" if (u or b) else ""
        if not block:
            continue
        token_count += _estimate_tokens(block)
        if token_count > MAX_CONTEXT_TOKENS:
            break
        lines.insert(0, block)
        
    recent_block = "\n\n".join(lines) if lines else "No recent turns."
    
    if compressed_context:
        return f"[Older Summary]\n{compressed_context}\n\n[Recent Turns]\n{recent_block}"
    
    return recent_block


def _truncate_history(history: list):
    """Keeps only the last MAX_HISTORY_TURNS messages in history."""
    if len(history) > MAX_HISTORY_TURNS:
        history[:] = history[-MAX_HISTORY_TURNS:]

# -----------------------------
# Configuration (from env)
# -----------------------------
AZ_SEARCH_ENDPOINT = os.getenv("AZ_SEARCH_ENDPOINT")
AZ_SEARCH_KEY = os.getenv("AZ_SEARCH_KEY")
AZ_SEARCH_INDEX = os.getenv("AZ_SEARCH_INDEX")
DRIVE_ID = os.getenv("DRIVE_ID")

AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
OPENAI_CHAT_DEPLOYMENT = os.getenv("OPENAI_CHAT_DEPLOYMENT")
OPENAI_EMBEDDING_DEPLOYMENT = os.getenv("OPENAI_EMBEDDING_DEPLOYMENT")
OPENAI_API_VERSION = os.getenv("OPENAI_API_VERSION", "2024-02-15-preview")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

GENERATED_DIR = os.getenv(
    "GENERATED_OUT_DIR",
    os.path.join(BASE_DIR, "generated_docs")
)
DEFAULT_CACHE_DIR = os.path.join(BASE_DIR, "generated_docs", "ppt_cache")
PPT_CACHE_DIR = os.getenv("GENERATED_PPT_CACHE_DIR", DEFAULT_CACHE_DIR)

# Ensure directories exist
os.makedirs(GENERATED_DIR, exist_ok=True)
os.makedirs(PPT_CACHE_DIR, exist_ok=True)

FILE_PATH = os.getenv("FILE_PATH")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINTURL = os.getenv("SHAREPOINTURL")
FOLDER_ID = os.getenv("FOLDER_ID")

# -----------------------------
# Initialize Azure clients
# -----------------------------
required_azure = [
    AZ_SEARCH_ENDPOINT,
    AZ_SEARCH_KEY,
    AZ_SEARCH_INDEX,
    AZURE_OPENAI_ENDPOINT,
    AZURE_OPENAI_API_KEY,
    OPENAI_CHAT_DEPLOYMENT,
    OPENAI_EMBEDDING_DEPLOYMENT
]

if not all(required_azure):
    raise ValueError("Missing critical Azure configuration in environment variables.")

try:
    embeddings_function = AzureOpenAIEmbeddings(
        azure_deployment=OPENAI_EMBEDDING_DEPLOYMENT,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_key=AZURE_OPENAI_API_KEY,
        api_version=OPENAI_API_VERSION
    ).embed_query

    vectorstore = AzureSearch(
        azure_search_endpoint=AZ_SEARCH_ENDPOINT,
        azure_search_key=AZ_SEARCH_KEY,
        index_name=AZ_SEARCH_INDEX,
        embedding_function=embeddings_function
    )

    llm = AzureChatOpenAI(
        azure_deployment=OPENAI_CHAT_DEPLOYMENT,
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version=OPENAI_API_VERSION,
        temperature=0,
        max_retries=3,
        verbose=True
    )

    classifier_llm = AzureChatOpenAI(
        azure_deployment=OPENAI_CHAT_DEPLOYMENT,
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version=OPENAI_API_VERSION,
        temperature=0,
        max_retries=3,
        verbose=True
    )

    logger.info("Successfully initialized Azure Search and OpenAI clients.")

except Exception as e:
    logger.error("Failed to initialize Azure clients: %s", e, exc_info=True)
    raise

# -----------------------------
# Session Store
# -----------------------------
_SESSIONS: Dict[str, Dict[str, Any]] = {}
_sessions_lock = threading.Lock()

def _get_or_create_session(session_id: Optional[str]) -> Dict[str, Any]:
    """Creates or retrieves an existing conversation session."""
    with _sessions_lock:
        if session_id not in _SESSIONS:
            # If no session_id provided, generate a new UUID for this session
            if not session_id:
                session_id = str(uuid.uuid4())

            # Create a new session if it doesn't already exist
            if session_id not in _SESSIONS:
                logger.info("Creating new session state for ID: %s", session_id)

                # Initialize all workflow fields to None based on the known schema
                workflow_state = {field: None for field in PITCH_DECK_FIELDS}
       
                # Add control variables that drive the conversation/workflow
                workflow_state.update({
                    "workflow": None,              # Name/identifier of the active workflow (if any)
                    "summary": None,               # Final/rolling summary text
                    "phase": "start",              # current phase of the flow
                    "last_question_asked_for": None,  # Tracks the last prompt key we asked the user
                    "retries": 0,                  # Track retries for smart fallbacks
                    "slide_count": None,           # Desired slide count, if collected/decided
                    "session_id": session_id,      # Echo the session ID inside the state for traceability
                    "compressed_context": ""       # Rolling memory for long conversations
                })

                # Create the session entry with empty history and initialized state
                _SESSIONS[session_id] = {
                    "history": [],     # conversation turns live here
                    "state": workflow_state
                }

    # Return the existing or newly created session object
    return _SESSIONS[session_id]

def _compress_history_if_needed(session: dict):
    """
    Rolling compression: if history exceeds the context window, summarize the 
    oldest turns into 'compressed_context' and remove them from the active list.
    """
    history = session["history"]
    state = session["state"]
    
    if len(history) > MAX_HISTORY_TURNS:
        # Keep the most recent MAX_CONTEXT_TURNS, compress the rest
        turns_to_compress = history[:-MAX_CONTEXT_TURNS]
        recent_to_keep = history[-MAX_CONTEXT_TURNS:]
        
        # Build text representation of turns to compress
        text_to_compress = "\n".join(
            f"User: {t.get('user','')}\nGia: {t.get('bot','')}" for t in turns_to_compress
        )
        old_summary = state.get("compressed_context", "")
        
        prompt = _compress_history_prompt(old_summary, text_to_compress)
        try:
            full_prompt = f"{MASTER_PROMPT}\n\n{prompt}"
            resp = llm.invoke(full_prompt)
            new_summary = resp.content if hasattr(resp, "content") else str(resp)
            if new_summary:
                state["compressed_context"] = new_summary
                logger.info("Successfully compressed older conversation history.")
        except Exception as e:
            logger.warning("Failed to compress history: %s", e)
            
        # Truncate the history array to just the recent turns
        session["history"] = recent_to_keep

# -----------------------------
# Safe LLM wrapper + Memory
# -----------------------------

def _safe_invoke_llm(prompt_text: str, chat_history = None, history: list = None, compressed_context: str = "") -> str:
    """
    Stateless LLM invocation — production-safe for multi-user environments.

    No shared global memory. Builds a per-call context block from the
    session's own history list, then calls the LLM directly.

    Args:
        prompt_text: The task-specific prompt to send to the LLM.
        history:     Optional session history list for context injection.
                     Uses _build_context_block() which is token-guarded.
        compressed_context: Optional rolling summary of older turns.
    """
    try:
        context_block = ""
        if history or compressed_context:
            ctx = _build_context_block(history, max_turns=MAX_CONTEXT_TURNS, compressed_context=compressed_context)
            if ctx and ctx != "No recent turns.":
                context_block = f"\n\n[Conversation so far]\n{ctx}"
        
        if chat_history:
            full_prompt = extract_CLEAN_RAG_PROMPT(context=chat_history,question=prompt_text)
            print(full_prompt)
        else:
            full_prompt = f"{MASTER_PROMPT}{context_block}\n\n{prompt_text}"
        resp = llm.invoke(full_prompt)
        return resp.content if hasattr(resp, "content") else str(resp)

    except Exception as e:
        logger.error("LLM failure in invocation (safe memory): %s", e)
        return "I encountered an error."

# sets loading status start

# deck_generation = False
# def says_is_loading_true_or_false(session) -> bool:
#     with _sessions_lock:
#         global deck_generation
#         if deck_generation:
#             return True
#         return False
    
# def set_deck_generation_to_true():
#     with _sessions_lock:
#         global deck_generation
#         deck_generation = True

# def set_deck_generation_to_false():
#     with _sessions_lock:
#         global deck_generation
#         deck_generation = False

# def check_deck_generation_status() -> bool:
#     with _sessions_lock:
#         global deck_generation
#         return deck_generation

deck_generation_status = {}

def says_is_loading_true_or_false(session_id) -> bool:
    if not session_id:
        # print("session_id_missing: ",session_id,deck_generation_status)
        return False
    if session_id not in deck_generation_status:
        # print("session_id missing from state",session_id,deck_generation_status)
        deck_generation_status[session_id] = False
    # print("session working fine",session_id,deck_generation_status)
    return deck_generation_status[session_id]

def set_deck_generation_to_true(session_id):
    with _sessions_lock:
        deck_generation_status[session_id] = True

def set_deck_generation_to_false(session_id):
    with _sessions_lock:
        deck_generation_status[session_id] = False

# sets loading status end

def generate_conversational_reply(situation: str, state: dict, history: list) -> str:
    """
    Core engine for dynamic LLM-generated conversational turns.
    Replaces get_dynamic_response templates with pure, situational generation.
    """
    comp_ctx = state.get("compressed_context", "") if isinstance(state, dict) else ""
    history_brief = _build_context_block(history, max_turns=MAX_CONTEXT_TURNS, compressed_context=comp_ctx)
    prompt = _generate_situational_response_prompt(situation, history_brief)
    
    try:
        resp = classifier_llm.invoke(prompt)
        return str(resp.content if hasattr(resp, "content") else resp).strip()
    except Exception as e:
        logger.error("LLM failure in conversational reply generator: %s", e)
        return "Something went wrong on my end, could you repeat that?"


# -----------------------------
# Intent + Smalltalk Classification Utilities
# -----------------------------


def llm_intent_classify(user_message: str, instruction: str) -> bool:
    """
    Robust LLM-based intent classifier.
    Returns True when the LLM answers 'yes' (semantically), False otherwise.
    This function is defensive: it sanitizes inputs and never raises.
    """
    try:
        # Normalize inputs: coerce None to empty strings and ensure str type
        um = "" if user_message is None else str(user_message)
        instr = "" if instruction is None else str(instruction)

        # Build a prompt for the LLM.
        # json.dumps is used to safely escape any special characters/newlines/quotes.
        # Invoke the classifier LLM
        resp_obj = classifier_llm.invoke(llm_intent_classify_prompt(instr=instr,um=um))
        raw = resp_obj.content if hasattr(resp_obj, "content") else str(resp_obj)

        # If invocation failed or returned empty, treat as no intent
        if not raw:
            return False

        # Normalize the model's response to lowercase and trim whitespace
        resp = str(raw).strip().lower()

        # Robust check for 'yes' or 'true' even if the LLM is verbose
        if "yes" in resp and "no" not in resp:
            return True
        if "true" in resp and "false" not in resp:
            return True
            
        # Fallback: Consider only the first token (handles cases like "yes.", "no, actually")
        first = resp.split()[0] if resp.split() else ""
        return first.startswith("y") or first.startswith("t")

    except Exception:
        # Defensive fallback: on any unexpected error, return False (no intent)
        return False


def is_no_value_answer_semantic(user_message: str) -> bool:
    """
    Uses the LLM classifier to decide if the user is intentionally indicating
    they have NO answer for a requested field (should be treated as a valid skip).
    """
    return llm_intent_classify(user_message, is_no_value_answer_semantic_prompt)


def is_smalltalk(user_message: str) -> bool:
    """
    Pure LLM-based smalltalk detector.
    Detects greetings, pleasantries, acknowledgements, chit-chat,
    or emotional check-ins with no actionable intent.
    """
    return llm_intent_classify(user_message, is_smalltalk_prompt)

def is_out_of_scope(user_message: str) -> bool:
    """
    Detects valid questions that are outside the pitch deck
    or business-document domain.
    """
    return llm_intent_classify(user_message, is_out_of_scope_prompt)


# -----------------------------
# Summary Generator
# -----------------------------
def build_context_profile(state: dict, summary_text: str) -> dict:
    """
    Builds a controlled context profile from user inputs.
    Works for both short prompts and long workflows.
    """

    text = (summary_text or "").lower()

    context = {
        "capability": state.get("capability"),
        "industry": None,
        "use_cases": set(),
        "audience": None,
        "solution_type": None
    }

    # -------------------
    # Industry detection
    # -------------------
    if any(k in text for k in ["loan", "credit", "bank", "financial"]):
        context["industry"] = "financial_services"
    elif "healthcare" in text:
        context["industry"] = "healthcare"
    elif any(k in text for k in ["retail", "ecommerce"]):
        context["industry"] = "retail"

    # -------------------
    # Use cases
    # -------------------
    if any(k in text for k in ["funnel", "drop-off", "conversion"]):
        context["use_cases"].add("funnel_analytics")

    if any(k in text for k in ["a/b", "experiment", "testing"]):
        context["use_cases"].add("ab_testing")

    if any(k in text for k in ["dashboard", "kpi", "metrics"]):
        context["use_cases"].add("reporting")

    if any(k in text for k in ["real-time", "near real time"]):
        context["use_cases"].add("real_time")

    # -------------------
    # Audience
    # -------------------
    if "marketing" in text:
        context["audience"] = "marketing"
    elif "product" in text:
        context["audience"] = "product"
    elif any(k in text for k in ["executive", "leadership", "cxo"]):
        context["audience"] = "leadership"

    # -------------------
    # Solution intent
    # -------------------
    if "proposal" in text:
        context["solution_type"] = "solution_proposal"
    elif "architecture" in text:
        context["solution_type"] = "architecture"

    return context



def generate_summary(answers: dict, workflow_name: str) -> str:
    """
    Builds a professional summary before deck generation.
    """
    # Validate workflow name against configured question sets
    if workflow_name not in WORKFLOW_QUESTIONS:
        return "Error: Invalid workflow."

    # Map each PITCH_DECK_FIELDS key to the corresponding workflow question
    question_map = dict(zip(PITCH_DECK_FIELDS, WORKFLOW_QUESTIONS.get(workflow_name, [])))

    # Collect lines for the "Client Summary" section
    client_summary_points = ["**Client Summary**", ""]

    # Build the list of questions for this workflow + the slide count question
    all_questions = WORKFLOW_QUESTIONS.get(workflow_name, []) + [SUMMARY_SLIDE_COUNT_QUESTION]

    # Iterate through all questions and add available answers to the summary
    for question in all_questions:
        answer = answers.get(question)
        if answer is not None:
            # Normalize label: strip '?', fix encoding, trim spaces
            label = question.replace("?", "").replace("â€™", "'").strip()

            # Apply friendly labels for specific questions
            if question.startswith("Which primary capability"):
                label = SUMMARY_CAPABILITY_LABEL
            elif question.startswith("Roughly how many slides"):
                label = SUMMARY_SLIDE_COUNT_LABEL

            # Append a bullet point for each answered question
            client_summary_points.append(f"- {label}: {answer}\n")

    # Join the client summary section into a single string
    client_summary = "\n".join(client_summary_points)

    return client_summary.strip()


def format_pitch_deck_summary_confirmation(state: dict) -> str:
    """
    Deterministic, structured confirmation message (reliable newlines/bullets).
    """
    slide_count = state.get("slide_count")
    bullets = []

    # Required fields
    if state.get("capability") not in (None, ""):
        bullets.append(f"- {SUMMARY_CAPABILITY_LABEL}: {state.get('capability')}")
    if state.get("client") not in (None, ""):
        bullets.append(f"- Client: {state.get('client')}")
    if state.get("challenge") not in (None, ""):
        bullets.append(f"- Challenge: {state.get('challenge')}")

    # Optional fields
    for field, label in [
        ("proposal", "Proposal"),
        ("experience", "Relevant Experience"),
        ("value", "Business Value"),
        ("offerings", "Offerings Supporting This"),
        ("audience_tone", "Audience & Tone"),
        ("deck_type", "Deck Type"),
    ]:
        v = state.get(field)
        if v not in (None, ""):
            bullets.append(f"- {label}: {v}")

    if slide_count is not None:
        bullets.append(f"- {SUMMARY_SLIDE_COUNT_LABEL}: {slide_count} slides")

    bullets_text = "\n".join(bullets) if bullets else "- (No details provided yet)"

    return (
        "Here's a quick summary of your project details:\n"
        f"{bullets_text}\n\n"
        "Does everything look accurate, or would you like to update any details before I start building your deck?\n"
        "Or, if everything looks good, just say 'ready' and I'll start building your deck!"
    )

# -----------------------------
# LLM-Based Field Extraction
# -----------------------------


def _extract_details_from_message(message: str, state: dict, history: list) -> dict:
    """
    LLM-driven structured extractor with:
    - Per-field semantic skip detection
    - Fallback skip detection when extraction is null
    - Field-specific context evaluation
    - Correction handling
    - Capability name/number resolution
    """

    message = (message or "").strip()
    if not message:
        return {}

    lower_msg = message.lower()

    # Detect if the user is correcting a previous answer
    is_correction = any(
        kw in lower_msg for kw in ["actually", "instead", "correction", "update", "change"]
    )

    # Determine fields to ask the LLM to extract
    missing_fields = [f for f in PITCH_DECK_FIELDS if state.get(f) is None]
    
    # If we are in the confirming phase or the user used correction keywords,
    # we must check ALL fields to catch their updates.
    in_confirm_phase = state.get("phase") == "confirming_assembly"
    if is_correction or in_confirm_phase or not missing_fields:
        fields_to_extract = PITCH_DECK_FIELDS + ["slide_count"]
    else:
        fields_to_extract = missing_fields

    if not fields_to_extract:
        return {}

    # Rich, token-guarded conversation context window
    comp_ctx = state.get("compressed_context", "") if isinstance(state, dict) else ""
    history_brief = _build_context_block(history, max_turns=MAX_CONTEXT_TURNS, compressed_context=comp_ctx)

    # -------------------------------
    # PRIMARY EXTRACTION PROMPT
    # -------------------------------
    raw = _safe_invoke_llm(
        _extract_details_from_message_prompt(
            history_brief=history_brief,
            message=message,
            missing_fields=fields_to_extract
        ),
        history=history,
        compressed_context=comp_ctx
    )
    json_match = re.search(r"\{[\s\S]*\}", raw)
    if not json_match:
        return {}

    try:
        parsed = json.loads(json_match.group(0))
    except Exception:
        return {}

    cleaned = {}

    # -------------------------------
    # PROCESS EACH FIELD
    # -------------------------------
    for field, extracted_value in parsed.items():

        if field not in fields_to_extract:
            continue

        # -------------------------------
        # CASE 1 - Extractor returned null
        # -------------------------------
        if extracted_value is None:
            # Local fast check: if the entire message was simply "skip", etc., mark all missing as skipped
            if _is_negative_skip_message(message):
                cleaned[field] = ""     # user intentionally skipped
            continue

        # Normal extracted string
        extracted_str = str(extracted_value).strip()

        # -------------------------------
        # CASE 2 - Per-field skip detection using extracted value
        # -------------------------------
        if extracted_str == "" or _is_negative_skip_message(extracted_str):
            cleaned[field] = ""
            continue

        # -------------------------------
        # Correction protection
        # -------------------------------
        existing = state.get(field)
        # In confirm phase, any extracted value is treated as a purposeful correction
        if existing not in (None, "") and not (is_correction or in_confirm_phase):
            # do not overwrite unless user explicitly corrects
            continue

        # -------------------------------
        # Capability handling
        # -------------------------------
        if field == "capability":
            chosen_cap = None

            # Numeric selection
            try:
                idx = int(extracted_str)
                if 1 <= idx <= len(AVAILABLE_CAPABILITIES):
                    chosen_cap = AVAILABLE_CAPABILITIES[idx - 1]
            except Exception:
                pass

            # Name selection
            if not chosen_cap:
                chosen_cap = next(
                    (c for c in AVAILABLE_CAPABILITIES
                     if c.lower() == extracted_str.lower()),
                    None
                )

            if chosen_cap:
                cleaned[field] = chosen_cap
            else:
                # User intentionally skipped capability
                if extracted_str == "":
                    cleaned[field] = ""
            continue

        # -------------------------------
        # Regular field assignment
        # -------------------------------
        cleaned[field] = extracted_str

    return cleaned


def get_next_question_from_llm(state: dict, history: list) -> Optional[str]:
    """
    Ask the next missing field question.
    Treat empty string "" as a valid answered field.
    Only fields with value = None are considered missing.
    """

    # Find next field whose value is None (not empty string)
    next_field = next((f for f in PITCH_DECK_FIELDS if state.get(f) is None), None)
    if not next_field:
        return None

    collected_info = "\n".join(
        f"- {k}: {v if v not in (None, '') else '(not provided)'}"
        for k, v in state.items()
        if k in PITCH_DECK_FIELDS
    )

    missing_summary = ", ".join(
        f for f in PITCH_DECK_FIELDS if state.get(f) is None
    )

    # Track retries for smart fallback
    if state.get("last_question_asked_for") == next_field:
        state["retries"] = state.get("retries", 0) + 1
    else:
        state["last_question_asked_for"] = next_field
        state["retries"] = 0

    retries = state.get("retries", 0)

    # Rich, token-guarded conversation context window
    comp_ctx = state.get("compressed_context", "") if isinstance(state, dict) else ""
    history_brief = _build_context_block(history, max_turns=MAX_CONTEXT_TURNS, compressed_context=comp_ctx)

    q = _safe_invoke_llm(
        _get_next_question_from_llm_prompt(
            collected_info=collected_info,
            history_brief=history_brief,
            missing_summary=missing_summary,
            next_field=next_field,
            retries=retries
        ),
        history=history,
        compressed_context=comp_ctx
    ).strip()

    # Add capability menu only when needed
    if next_field == "capability":
        q += f"\n\nPlease choose one of the following options:\n{CAPABILITY_OPTIONS_STR}"

    state["last_question_asked_for"] = next_field
    return q

def _extract_control_json_from_text(text: str) -> Optional[dict]:
    """Extracts a control JSON block if it exists."""
    if not text:
        return None

    m = CONTROL_JSON_REGEX.search(text)
    if not m:
        return None

    try:
        return json.loads(m.group(1))
    except Exception:
        return None


def _validate_generate_payload(payload: dict) -> Tuple[bool, str]:
    """Ensures JSON for deck generation is valid."""
    if not isinstance(payload, dict):
        return False, "Payload must be an object."

    if payload.get("workflow") != "pitch_deck":
        return False, "Invalid workflow."

    if not isinstance(payload.get("fields", {}), dict):
        return False, "Fields must be a JSON object."

    add_slides = payload.get("additional_slides")
    if add_slides is None:
        return False, "Missing additional_slides."

    try:
        int(add_slides)
    except Exception:
        return False, "additional_slides must be an integer."

    return True, ""

# -----------------------------
# Ask LLM for GENERATE_PPT Command
# -----------------------------


def ask_llm_for_generate_control(summary_text: str, current_state: dict) -> Tuple[Optional[dict], str]:
    """
    Smart context gate: the LLM reviews the full summary + state and ONLY emits
    a GENERATE_PPT JSON block when it is fully confident everything is correct.
    If anything is unclear or missing, it returns a clarifying question instead.

    Returns:
        (control_dict, raw_text)
        - control_dict is the validated JSON payload, or None if not ready.
        - raw_text is the full LLM response (use as clarifying question when control_dict is None).
    """
    raw = _safe_invoke_llm(
        _ask_llm_for_generate_control_prompt(
            summary_text=summary_text,
            current_state=current_state
        )
    )

    control = _extract_control_json_from_text(raw)

    if not control or control.get("action") != "GENERATE_PPT":
        return None, raw  # raw contains the LLM's clarifying question

    ok, err = _validate_generate_payload(control.get("payload", {}))
    if not ok:
        logger.warning("GENERATE_PPT payload validation failed: %s", err)
        return None, raw

    return control, raw

def is_custom_slide_request(message: str) -> bool:
    return llm_intent_classify(message, is_custom_slide_request_prompt)

def extract_custom_slide_request(message: str) -> dict:
    raw = _safe_invoke_llm(_extract_custom_slide_request_prompt(message=message))

    try:
        json_match = re.search(r"\{[\s\S]*\}", raw)
        return json.loads(json_match.group(0))
    except:
        return {"slides": [1], "topics": [message.strip()]}

# def assume_custom_slides_based_on_user_input(extracted_slides, user_input) -> List[int]:
    # print("extracted_slides",extracted_slides)
    # # Case 1: user_input == 0 → replace all with 0
    # if user_input == 0:
    #     print("input == 0 ",[0 for _ in extracted_slides], type([0 for _ in extracted_slides]))
    #     return [0 for _ in extracted_slides]

    # total = sum(extracted_slides)

    # # Case 2: user_input < total but > 0 → replace all with 1
    # if 0 < user_input < total:
    #     print("input < total ",[1 for _ in extracted_slides], type([1 for _ in extracted_slides]))
    #     return [1 for _ in extracted_slides]


    # Case 3: user_input > total → increment until sum matches user_input
    # result = extracted_slides[:]
    # while sum(result) < user_input:
    #     for i in range(len(result)):
    #         if result[i] < user_input and sum(result) < user_input:
    #             result[i] += 1
    # print("sum(result) < user_input", result, type(result))
    # return result
    
def generate_custom_ppt(topics: List[str], slides: List[int], state: Dict):
    session_id = state.get("session_id")
    set_deck_generation_to_true(session_id=session_id)
    try:
        token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)

        if DRIVE_ID:
            sp_drive_id = DRIVE_ID
        else:
            site_id, sp_drive_id = get_site_and_drive_id(token, SHAREPOINTURL)

        final_slides = []
        topic_append = []
        total_slides = 0
        exclude_files = []
        # Collect slides for each topic
        for topic, slide_count in zip(topics, slides):
            slides_for_topic = agent_tools.find_relevant_slides(
                endpoint=AZ_SEARCH_ENDPOINT,
                key=AZ_SEARCH_KEY,
                index_name=AZ_SEARCH_INDEX,
                openai_endpoint=AZURE_OPENAI_ENDPOINT,
                openai_key=AZURE_OPENAI_API_KEY,
                openai_deployment=OPENAI_EMBEDDING_DEPLOYMENT,
                chat_deployment="gpt-4.1",
                query=topic,
                num_slides=slide_count,
                cache_dir=PPT_CACHE_DIR,
                token=token,
                drive_id=sp_drive_id,
                capability=None,
                context_profile=None,
                exclude_files_content=None#exclude_files
            )
            # exclude_files.extend(slides_for_topic.get("exclude_files_content", []))
            final_slides.extend(slides_for_topic.get("final_slides", []))
            topic_append.append(topic)
            print("topic",topic)
            print("slide_count",slide_count)
            total_slides += slide_count
        exclude_files.clear()
        # Build final deck
        joined_topic = " ".join(topic_append)
        filename = f"Custom_{joined_topic[:40]}_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.pptx"
        output_path = os.path.join(GENERATED_DIR, filename)

        final_path = agent_tools.create_dynamic_pitch_deck(
                initial_slides=final_slides,
                additional_slides=[],
                summary_text=f"Custom deck on {joined_topic}",
                output_path=output_path,
                cache_dir=PPT_CACHE_DIR
        )
        
        state["workflow"] = None
        state["phase"] = "start"
        set_deck_generation_to_false(session_id=session_id)
        if not final_path:
            logger.error("error in dynamic pitch deck generation")
        return {
                "answer": f"Your {total_slides}-slide presentation on custom requirements is now ready.",
                "file": filename
            }

    except Exception as e:
            set_deck_generation_to_false(session_id=session_id)
            logger.error("Custom PPT generation failed: %s", e, exc_info=True)
            return {"answer": "Failed to generate custom slides."}

# -----------------------------
# Pitch Deck Workflow Handler
# -----------------------------

def handle_pitch_deck_workflow(message: str, state: dict, history: list) -> dict:
    """
    Main logic for the pitch deck workflow.
    Handles phases:
      1. gathering
      2. awaiting_slide_count
      3. confirming_assembly
      4. generating
    """
    session_id = state.get("session_id")


    # -----------------------------------------------------
    # PHASE: Custom — Ask for slide counts for missing topics
    # -----------------------------------------------------
    if state.get("phase") == "ask_slides_gathering_permission":
        # num_to_match = llm.invoke(_ask_slides_gathering_permission_reply_prompt(message))
        # try:
        #     json_match = re.search(r"\{[\s\S]*\}", num_to_match)
        #     num_match = json.loads(json_match.group(0))
        # except:
        #     return {"slides": [1], "topics": [message.strip()]}
        # if num_to_match.get("count") == 1:
        #     return generate_custom_ppt(
        #         slides=[1 if x == 0 else x for x in state.get("extracted_slides_counts")],
        #         state=state,
        #         topics=state.get("extracted_topics")
        #     )
        # elif num_to_match.get("count") == 2:
        #     return generate_custom_ppt(
        #         slides=num_to_match.get("slides"),
        #         state=state,
        #         topics=state.get("extracted_topics")
        #     )
        # else :
        #     state["phase"] = "slides_gathering"
        #     return {
        #         "answer": f"The deck generation is now underway. Please specify the number of slides you would like for the deck ?"
        #     }
        try:
            # invoke the LLM
            if "skip" in message:
                    state["workflow"] = None
                    state["phase"] = "start"
                    return {"answer" : "deck generation cancelled."}
                
            raw_msg = llm.invoke(_ask_slides_gathering_permission_reply_prompt(message))
            print(raw_msg, type(raw_msg))

            # extract the text content from the AIMessage
            raw = raw_msg.content if hasattr(raw_msg, "content") else str(raw_msg)

            # extract JSON block if extra text is present
            json_match = re.search(r"\{[\s\S]*\}", raw)
            if not json_match:
                raise ValueError("No JSON object found in LLM output")

            num_to_match = json.loads(json_match.group(0))

            # enforce integer types
            slides_val = int(num_to_match.get("slides", 0))
            count_val = int(num_to_match.get("count", 3))
            num_to_match = {"slides": slides_val, "count": count_val}

        except Exception as e:
            print("parsing failed", e)
            # fallback
            return {"slides": [1], "topics": [message.strip()]}

        # Case 1: positive affirmation → proceed with detected slides
        if num_to_match["count"] == 1:
            return generate_custom_ppt(
                slides=[1 if x == 0 else x for x in state.get("extracted_slides_counts")],
                state=state,
                topics=state.get("extracted_topics")
            )
        # Case 2: negative + number → use custom slide count
        elif num_to_match["count"] == 2:
            return generate_custom_ppt(
                slides=[num_to_match["slides"]],
                state=state,
                topics=state.get("extracted_topics")
            )
        # Case 3: ambiguous → ask user to specify
        else:
            state["phase"] = "slides_gathering"
            # answers = llm.invoke(_deck_generation_status_decline_prompt(message))
            # return {
            #     "answer": answers.content
            # }
            return {
                "answer": "The deck generation is now underway. Please specify the number of slides you would like for the deck."
            }



    if state.get("phase") == "slides_gathering":
            try:
                chat = message
                if "skip" in chat:
                    state["workflow"] = None
                    state["phase"] = "start"
                    return {"answer" : "deck generation cancelled."}
                num = int(message)
                return generate_custom_ppt(
                        # slides=assume_custom_slides_based_on_user_input(extracted_slides=state.get("extracted_slides_counts", []),user_input=num),
                        # slides=num,
                        slides=[num],
                        state=state,
                        topics=state.get("extracted_topics")
                    )

            except ValueError:
                return {
                    "answer": "Please provide a valid number of slides."
                }
            except Exception as e:
                logger.error("error in handle_pitch_deck_workflow(): ",e)
                return{
                    "answer": "Something went wrong...!"
                }

    # -----------------------------------------------------
    # PHASE 1 — GATHER INFORMATION
    # -----------------------------------------------------
    if state.get("phase") == "gathering":
        if message:
            # Loop guard: if we already asked for a specific field and the user
            # replies with a bare "no/skip", treat it as an explicit skip.
            last_field = state.get("last_question_asked_for")
            handled_no_answer = False
            if (
                last_field in PITCH_DECK_FIELDS
                and state.get(last_field) is None
                and _is_negative_skip_message(message)
            ):
                state[last_field] = ""
                handled_no_answer = True
                logger.info("Loop guard: user skipped '%s' in gathering phase.", last_field)

            # Extract structured fields (e.g., client, capability, challenge) from free-text
            extracted = {} if handled_no_answer else _extract_details_from_message(message, state, history)

            if extracted:
                logger.info("Extracted fields: %s", extracted)

                for field, value in extracted.items():
                    if value is None:
                        continue  # Skip empty/None values from extraction

                    # Determine if the user is correcting a previous answer
                    msg_lower = (message or "").lower()
                    is_correction = any(
                        kw in msg_lower
                        for kw in ["actually", "instead", "change", "update", "correction"]
                    )

                    existing = state.get(field)
                    # If a value already exists and this isn't a correction, don't overwrite
                    if existing not in (None, "") and not is_correction:
                        continue

                    # Special handling for 'capability' to support numeric choice or exact text match
                    if field == "capability":
                        chosen_cap = None
                        try:
                            # If user provided a number (1-based), map it to AVAILABLE_CAPABILITIES
                            idx = int(str(value).strip())
                            if 1 <= idx <= len(AVAILABLE_CAPABILITIES):
                                chosen_cap = AVAILABLE_CAPABILITIES[idx - 1]
                        except Exception:
                            pass

                        # Fallback: try exact (case-insensitive) text match
                        if not chosen_cap:
                            chosen_cap = next(
                                (c for c in AVAILABLE_CAPABILITIES
                                 if c.lower() == str(value).strip().lower()),
                                None
                            )

                        if chosen_cap is not None:
                            state[field] = chosen_cap
                        else:
                            # Allow empty-string capability, but warn on invalid non-empty input
                            if str(value).strip() == "":
                                state[field] = ""
                            else:
                                logger.warning("Invalid capability extracted: %s", value)
                        continue  # Capability handled; move to next extracted field

                    # Default path: set the extracted value directly into state
                    state[field] = value

        # Ask the next best question (driven by LLM) if there are still gaps
        next_q = get_next_question_from_llm(state, history)
        if next_q:
            history.append({"user": message, "bot": next_q})
            return {"answer": next_q}

        # If no more questions to ask, transition to requesting slide count
        state["phase"] = "awaiting_slide_count"
        ask_slide_count_msg = generate_conversational_reply(
            "Acknowledge that you have captured all the necessary details. Ask the user how many slides they would like to append to the deck.",
            state, history
        )
        history.append({"user": message, "bot": ask_slide_count_msg})
        return {"answer": ask_slide_count_msg}

    # -----------------------------------------------------
    # PHASE 2 — EXPECTING SLIDE COUNT
    # -----------------------------------------------------
    if state.get("phase") == "awaiting_slide_count":
        try:
            # Parse the user's message as an integer and clamp to >= 0
            num = int(message)
            state["slide_count"] = max(0, num)
        except Exception:
            invalid_msg = generate_conversational_reply(
                "Tell the user they didn't provide a valid number. Ask them to reply with a valid integer for the slide count.",
                state, history
            )
            return {"answer": invalid_msg}

        # Build a map from field keys -> questions for this workflow
        question_map = dict(zip(PITCH_DECK_FIELDS, WORKFLOW_QUESTIONS["pitch_deck"]))
        # Gather all answered questions from state
        answers = {q: state.get(f) for f, q in question_map.items() if state.get(f) is not None}
        # Add the slide count answer into the answers used for summary
        answers["Roughly how many slides should the pitch deck be?"] = state["slide_count"]

        # Generate a professional summary to confirm scope before assembly
        state["summary"] = generate_summary(answers, "pitch_deck")
        state["phase"] = "confirming_assembly"
        # Present the compiled summary and ask for confirmation to proceed
        summary_msg = format_pitch_deck_summary_confirmation(state)
        history.append({"user": message, "bot": summary_msg})
        return {"answer": summary_msg}

    # -----------------------------------------------------
    # PHASE 3 — CONFIRMATION FOR GENERATION
    # -----------------------------------------------------
    if state.get("phase") == "confirming_assembly":

        # Loop guard for smart-gate clarifying questions:
        # If Gia previously asked for missing required info, remember which
        # required field she was asking for, and let bare "no/yes" responses
        # skip it instead of re-asking forever.
        pending_field = state.get("pending_clarification_field")
        pending_retry_count = state.get("smart_gate_retry_count", 0)
        if (
            pending_field in PITCH_DECK_FIELDS
            and state.get(pending_field) is None
            and _is_negative_skip_message(message)
        ):
            state[pending_field] = ""
            state.pop("pending_clarification_field", None)
            state.pop("smart_gate_retry_count", None)
            state.pop("last_smart_gate_field", None)
            logger.info("Loop guard: skipped pending field '%s' after user said no/skip.", pending_field)

            # Pro behavior: if the user says "no/skip" to a smart-gate clarifier,
            # stop looping and proceed with best-effort deck generation.
            # (The controller may keep asking because required details are blank.)
            question_map = dict(zip(PITCH_DECK_FIELDS, WORKFLOW_QUESTIONS["pitch_deck"]))
            answers = {q: state.get(f) for f, q in question_map.items() if state.get(f) is not None}
            answers["Roughly how many slides should the pitch deck be?"] = state.get("slide_count", 0)
            state["summary"] = generate_summary(answers, "pitch_deck")

            # Immediately trigger the generating phase and let this same call produce the deck.
            state["phase"] = "generating"
            return handle_pitch_deck_workflow(message, state, history)

        # Another loop-breaker: if Gia asked for something and the user only
        # said "yes" without providing details, don't re-ask endlessly.
        if (
            pending_field in PITCH_DECK_FIELDS
            and state.get(pending_field) is None
            and pending_retry_count >= 1
            and _is_affirmation_only_message(message)
        ):
            state[pending_field] = ""
            state.pop("pending_clarification_field", None)
            state.pop("smart_gate_retry_count", None)
            state.pop("last_smart_gate_field", None)
            logger.info("Loop guard: skipped pending field '%s' after repeated affirmation without details.", pending_field)

            control, raw_llm_resp = ask_llm_for_generate_control(
                summary_text=state.get("summary", ""),
                current_state=state
            )
            if control:
                payload_fields = control.get("payload", {}).get("fields", {})
                for f, v in payload_fields.items():
                    if f in PITCH_DECK_FIELDS and v and state.get(f) in (None, ""):
                        state[f] = v
                # Respect slide count from control payload if provided
                ctrl_slides = control.get("payload", {}).get("additional_slides")
                if ctrl_slides is not None:
                    try:
                        state["slide_count"] = max(0, int(ctrl_slides))
                    except Exception:
                        pass
                
                # Proceed directly to generation without asking further questions
                state["phase"] = "generating"
                return handle_pitch_deck_workflow(message, state, history)

            clarifying_q = (raw_llm_resp or "").strip()
            if not clarifying_q or len(clarifying_q) > 600:
                clarifying_q = generate_conversational_reply(
                    "Gently tell the user you want to make sure you have everything right before generating. Ask the single most important missing or unclear detail.",
                    state, history
                )
            state["pending_clarification_field"] = _get_missing_required_field(state)
            if state["pending_clarification_field"]:
                if state.get("last_smart_gate_field") == state["pending_clarification_field"]:
                    state["smart_gate_retry_count"] = state.get("smart_gate_retry_count", 0) + 1
                else:
                    state["smart_gate_retry_count"] = 1
                    state["last_smart_gate_field"] = state["pending_clarification_field"]
            history.append({"user": message, "bot": clarifying_q})
            return {"answer": clarifying_q}

        # ✅ PRIORITY 1: Check for corrections (slide count or field updates)
        made_corrections = False
        
        # 1a. Hardcoded regex check for slide count (fast and reliable)
        lower_msg = message.lower()
        if "slide" in lower_msg or "count" in lower_msg:
            num_match = re.search(r'\b(\d+)\b', lower_msg)
            if num_match:
                state["slide_count"] = max(0, int(num_match.group(1)))
                made_corrections = True
                logger.info("Hardcoded slide count updated to %s", state["slide_count"])

        # 1b. LLM extraction for text fields
        extracted = _extract_details_from_message(message, state, history)
        
        if extracted:
            for field, value in extracted.items():
                if value is None or value == "":
                    continue
                
                # Special handling for capability (matches gathering phase)
                if field == "capability":
                    chosen_cap = None
                    try:
                        idx = int(str(value).strip())
                        if 1 <= idx <= len(AVAILABLE_CAPABILITIES):
                            chosen_cap = AVAILABLE_CAPABILITIES[idx - 1]
                    except Exception:
                        pass
                    if not chosen_cap:
                        chosen_cap = next((c for c in AVAILABLE_CAPABILITIES if c.lower() == str(value).strip().lower()), None)
                    
                    if chosen_cap:
                        state[field] = chosen_cap
                        made_corrections = True
                elif field == "slide_count":
                    try:
                        # Extract first number from the extracted string
                        num_match = re.search(r'\d+', str(value).strip())
                        if num_match:
                            state["slide_count"] = max(0, int(num_match.group()))
                            made_corrections = True
                    except Exception:
                        pass
                else:
                    state[field] = value
                    made_corrections = True
                    
        if made_corrections:
            logger.info("Corrections detected in confirm phase. Regenerating summary.")
            question_map = dict(zip(PITCH_DECK_FIELDS, WORKFLOW_QUESTIONS["pitch_deck"]))
            answers = {q: state.get(f) for f, q in question_map.items() if state.get(f) is not None}
            answers["Roughly how many slides should the pitch deck be?"] = state.get("slide_count", 0)
            
            state["summary"] = generate_summary(answers, "pitch_deck")
            # Stay in confirming_assembly, but prompt with updated summary
            update_msg = generate_conversational_reply(
                f"Tell the user you have successfully applied their changes. Present the revised summary:\n\n{state['summary']}\n\nAsk them if they are ready to generate the deck now.",
                state, history
            )
            history.append({"user": message, "bot": update_msg})
            return {"answer": update_msg}

        # ✅ PRIORITY 2: Check confirmation BEFORE cancellation.
        # "No" as an answer to "would you like to tweak anything?" means proceed — not cancel.
        if llm_intent_classify(message, is_confirmation_prompt):
            logger.info("Session %s: User confirmed. Proceeding directly to generation.", session_id)
            
            # Clear any pending smart-gate clarification state
            state.pop("pending_clarification_field", None)
            state.pop("smart_gate_retry_count", None)
            state.pop("last_smart_gate_field", None)
            
            # Proceed directly to generation without asking further questions
            state["phase"] = "generating"
            return handle_pitch_deck_workflow(message, state, history)

        # ✅ PRIORITY 3: Only cancel if the user has NOT confirmed and is explicitly stopping.
        elif llm_intent_classify(message, is_cancel_deck_prompt):
            if session_id:
                _SESSIONS.pop(session_id, None)
            cancel_msg = generate_conversational_reply(
                "Acknowledge the cancellation and warmly tell the user the deck generation has been stopped.",
                state, history
            )
            history.append({"user": message, "bot": cancel_msg})
            return {"answer": cancel_msg}

        else:
            # Unclear — ask the user to clarify
            clarify_msg = generate_conversational_reply(
                "Gently tell the user you didn't quite catch their meaning. Ask them clearly if they want to generate the deck, cancel it, or update a detail.",
                state, history
            )
            history.append({"user": message, "bot": clarify_msg})
            return {"answer": clarify_msg}

    # -----------------------------------------------------
    # PHASE 3b — CONTEXT ECHO & SMART PRE-GENERATION CHECK
    # -----------------------------------------------------
    if state.get("phase") == "ready_to_generate":
        # The bot has echoed its understanding. Now:
        # - If user says 'yes' / 'sounds good' / 'go ahead' → generate
        # - If user corrects something → update state and ask again
        # - If user asks a clarifying question / adds detail → absorb it and generate

        # 1. Try to extract any last-minute corrections from the user's response
        extracted = _extract_details_from_message(message, state, history)
        if extracted:
            for field, value in extracted.items():
                if value is not None and value != "" and field in PITCH_DECK_FIELDS:
                    state[field] = value
                elif field == "slide_count" and value not in (None, ""):
                    try:
                        state["slide_count"] = max(0, int(str(value)))
                    except Exception:
                        pass

        # 2. Check if user wants to cancel
        if llm_intent_classify(message, is_cancel_deck_prompt):
            if session_id:
                _SESSIONS.pop(session_id, None)
            cancel_msg = generate_conversational_reply(
                "Acknowledge the cancellation and warmly tell the user the deck generation has been stopped.",
                state, history
            )
            history.append({"user": message, "bot": cancel_msg})
            return {"answer": cancel_msg}
        # 3. Move to actual generation — absorb any additions and proceed
        state["phase"] = "generating"
        assembling_msg = generate_conversational_reply(
            "Enthusiastically tell the user that you are now starting to build their deck. "
            "Mention 1-2 specific details from their input (client name, capability, or challenge) to show you fully understood them. "
            "Ask them to hold on for a moment while it generates.",
            state, history
        )
        history.append({"user": message, "bot": assembling_msg})
        # Fall through to Phase 4 below

    # -----------------------------------------------------
    # PHASE 4 — GENERATING FINAL PPT
    # -----------------------------------------------------
    if state.get("phase") == "generating":
        set_deck_generation_to_true(session_id=session_id)
        logger.info("Session %s: Beginning PPT generation", session_id)

        # 1. Build Context Profile from collected state and summary
        state["context_profile"] = build_context_profile(
            state,
            state.get("summary", "")
        )
        logger.info("Context profile created: %s", state["context_profile"])

        try:
            # Extract commonly used inputs from state
            capability = state.get("capability") or "Pitch"
            extra = int(state.get("slide_count") or 0)
            client = state.get("client") or "the client"
            challenge = state.get("challenge") or ""
           
            # --- AUTHENTICATION ---
            # Obtain OAuth token for Microsoft Graph
            token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
            
            # Resolve Drive ID: prefer configured DRIVE_ID, else fetch from the SharePoint site
            if DRIVE_ID:
                sp_drive_id = DRIVE_ID
            else:
                site_id, sp_drive_id = get_site_and_drive_id(token, SHAREPOINTURL)

            # --- SEARCH SLIDES ---
            logger.info("Calling find_relevant_slides...")
            slides_dict = agent_tools.find_relevant_slides(
                endpoint=AZ_SEARCH_ENDPOINT,
                key=AZ_SEARCH_KEY,
                index_name=AZ_SEARCH_INDEX,
                openai_endpoint=AZURE_OPENAI_ENDPOINT,
                openai_key=AZURE_OPENAI_API_KEY,
                openai_deployment=OPENAI_EMBEDDING_DEPLOYMENT,
                # 🟢 ADD THIS for "The Brain" (Query Expansion)
                # Ensure this variable matches your GPT-4 deployment name variable in this file
                chat_deployment="gpt-4.1",
                query=f"{capability} deck for {client} focusing on {challenge}",
                num_slides=extra,
                cache_dir=PPT_CACHE_DIR,
                token=token,
                drive_id=sp_drive_id,
                capability=capability,
                context_profile=state["context_profile"]
            )
            slides = slides_dict.get("final_slides",[])

            # --- ASSEMBLE DECK ---
            logger.info(f"Found {len(slides)} slides. Assembling deck...")
           
            # Compute a unique filename using session and timestamp // this block is commented because it creates downloadable file with random letters and number as the file_name
           # timestamp = int(datetime.now().timestamp())
           # filename = f"Pitch_{session_id}_{timestamp}.pptx"
           # output_path = os.path.join(GENERATED_DIR, filename)
           
           # --- NEW CLEAN FILENAME LOGIC ---
           # Grab the capability(DOMAIN) , fallback to "Pitch" if missing
            domain_name = capability if capability else "Pitch"
           # Make it safe for saving (remove spaces , eg ., "Cloud Services")
            safe_domain = re.sub(r'[^a-zA-Z0-9]','-',domain_name)
           #Grab the client_name 
            client_name = client if client else "Client" 
           # Make it safe for saving (remove spaces , eg ., "Cloud Services")
            safe_client = re.sub(r'[^a-zA-Z0-9]','-',client_name)
           # Create the readable timestamp(Format: DD-MM-YY_HH-MM-AM)
            readable_timestamp = datetime.now().strftime("%d-%m-%Y_%I-%M-%p")
           # Build the final filename
            filename = f"{safe_domain}_{safe_client}_{readable_timestamp}.pptx"
            output_path = os.path.join(GENERATED_DIR, filename)
           # --------------------------------------------------------------------

            # Build the final deck from selected slides
            final_path = agent_tools.create_dynamic_pitch_deck(
                initial_slides=slides,
                additional_slides=[],
                summary_text=state.get("summary", ""),
                output_path=output_path,
                cache_dir=PPT_CACHE_DIR
            )
            set_deck_generation_to_false(session_id=session_id)
            # Report success (with downloadable file) or a generic failure
            if final_path:
                msg = generate_conversational_reply(
                    "Excitedly tell the user that their pitch deck is completely ready and they can download it below.",
                    state, history
                )
                history.append({"user": message, "bot": msg})

                # ---------------------------------------------------------
                #  UPDATED: THE SESSION RESET LOGIC
                # Wipes the memory so the next "Hi" doesn't loop the code
                # ---------------------------------------------------------
                logger.info(f"DEBUG: Deck generated: {filename}. Resetting state to start.")
                state["workflow"] = None
                state["phase"] = "start"
                for field in PITCH_DECK_FIELDS:
                    state[field] = None
                
                #  NEW: Clear the remaining state and wipe the LLM's memory
                state["slide_count"] = None
                state["summary"] = None
                history.clear() 
                # ---------------------------------------------------------
                
                # SEND THE TIME TO APP.PY
                return {"answer": msg, "file": filename}
            else:
                state["workflow"] = None
                state["phase"] = "start"
                err_msg = generate_conversational_reply(
                    "Apologize to the user and tell them you ran into a snag while assembling the slides. Ask them to try again.",
                    state, history
                )
                return {"answer": err_msg}

        except Exception as e:
            set_deck_generation_to_false(session_id=session_id)
            logger.error("Deck generation failed: %s", e, exc_info=True)
            state["workflow"] = None
            state["phase"] = "start"
            err_msg = generate_conversational_reply(
                "Apologize deeply and tell the user an unexpected error occurred while generating the presentation. Ask them to try one more time.",
                state, history
            )
            return {"answer": err_msg}


# -----------------------------
# RAG Retrieval Function
# -----------------------------
def run_qa(query: str, chat_history: List[dict]) -> dict:
    """
    Hybrid Retrieval-Augmented Generation.
    Uses a CLEAN system prompt to avoid crashing on {variable} placeholders in the Master Prompt.
    """
    try:
        formatted_history = []

        for entry in chat_history[-MAX_HISTORY_TURNS:]:
                if entry["role"] == "user":
                    formatted_history.append("user :"+entry["content"])
                else:
                    formatted_history.append("ai :"+entry["content"])

        # Execute the chain with the current question and recent conversation history
        result = _safe_invoke_llm(prompt_text=query,chat_history=formatted_history)

        # Return whatever the chain provides (e.g., {"answer": ..., "source_documents": ...})
        return result

    except Exception as e:
        # Log the full stack trace and return a safe fallback response
        logger.error("RAG QA failed: %s", e, exc_info=True)
        return generate_conversational_reply("Inform the user you ran into an issue while retrieving information and ask them to try again.", state={}, history=chat_history)

# -----------------------------
# Main Chat Router
# -----------------------------

def chat_once(message: str, session_id: Optional[str] = None, chat_data: Optional[str] = None) -> dict:
    """
    Central router for ALL interactions.
    Enhanced so the very first user message is processed (intent / extraction)
    before returning the initial greeting. This prevents the greeting from
    swallowing the user's first actionable message.
    """

    # Track whether this invocation created a brand-new session (affects greeting later)
    is_new_session = False
    # Normalize message for lightweight keyword checks
    msg_lower = (message or "").lower().strip()

    # Create new session if no session_id was provided by the caller
    if not session_id:
        session_id = str(uuid.uuid4())
        is_new_session = True
        logger.info("Started NEW session: %s", session_id)

    # Retrieve (or create) the session container holding state + history
    session = _get_or_create_session(session_id)
    
    # 🟢 NEW: Trigger compression if history is getting too long to save tokens
    _compress_history_if_needed(session)
    
    state = session["state"]
    history = session["history"]
    # ---------------------------------------------
    # 🔄 Restart triggers (semantic)
    # ---------------------------------------------
    # If the user asks to restart, reset state by dropping current session
    if any(k in msg_lower for k in ["restart", "start over", "reset", "begin again", "clear"]) and llm_intent_classify(message, is_restart_prompt):
        logger.info("Restart requested for session %s", session_id)
        if session_id in _SESSIONS:
            _SESSIONS.pop(session_id, None)
        new_id = str(uuid.uuid4())
        restart_msg = generate_conversational_reply(
            "Tell the user you are starting fresh. Ask them if they'd like to create a Proposal Deal Deck or a Pitch Deck.",
            state, history
        )
        return {"answer": restart_msg, "session_id": new_id, "sources": []}
    
    # ---------------------------------------------
    # ⚡ CUSTOM SLIDE FAST PATH
    # ---------------------------------------------
    if "slide" in msg_lower and is_custom_slide_request(message):
        extracted = extract_custom_slide_request(message)
        slides = extracted.get("slides")
        topic = extracted.get("topics", message)
        missing_topics = extracted.get("missing_topics")

        if missing_topics:
            state["workflow"] = "custom_pitch_deck"
            state["phase"] = "ask_slides_gathering_permission"
            # state["extracted_slides_counts"] = slides
            # state["extracted_topics"] = topic
            state["extracted_slides_counts"] = [sum(slides)]
            state["extracted_topics"] = ["".join(topic)]
            logger.info("Custom slide request detected: %s slides on %s", slides, topic)
            # answer = llm.invoke(_missing_topic_response_prompt(slides = sum(slides), message= message))
            # return { "answer": answer.content}
            return { "answer": f"I have detected {sum(slides)} slides. If you’d like to continue with {sum(slides)} slides, just give me a positive affirmation, and I’ll start the deck generation with {sum(slides)} slides. Feel free to say no if you’d prefer custom slide count instead."}
        
        logger.info("Custom slide request detected: %s slides on %s", slides, topic)
        # result = generate_custom_ppt(topic, slides, state)
        result = generate_custom_ppt(["".join(topic)], [sum(slides)], state)
        result.setdefault("session_id", session_id)
        result.setdefault("sources", [])
        return result

    # ---------------------------------------------
    # 🧩 Custom Pitch Deck workflow
    # ---------------------------------------------
    if state.get("workflow") == "custom_pitch_deck":
        return handle_pitch_deck_workflow(message, state,history=history)

    # ---------------------------------------------
    # ✳️ Improved Intent Detection (runs even for first message)
    # ---------------------------------------------
    # High-level intent: does the user want a pitch/presentation deck?
    user_is_requesting_pitch_deck = False
    
    # If we're at the very start of the workflow, route based on the user's intent
    if state.get("phase") == "start":
        user_is_requesting_pitch_deck = llm_intent_classify(message, is_pitch_deck_request_prompt)
        
        # A. Direct ask for a pitch deck -> enter the pitch deck workflow
        if user_is_requesting_pitch_deck:
            state["workflow"] = "pitch_deck"
            state["phase"] = "gathering"
            logger.info("Pitch deck workflow triggered — session %s", session_id)

            # 👇 UPDATE 2: Pass 'message' and remove duplicate history append
            init_q = handle_pitch_deck_workflow(message, state, history)
            init_q.setdefault("session_id", session_id)
            init_q.setdefault("sources", [])
            return init_q

        # B. Generic verbs + "deck" without type -> ask user to choose the deck type
        if ("create" in msg_lower or "make" in msg_lower or "build" in msg_lower) and "deck" in msg_lower:
            deck_type_msg = generate_conversational_reply(
                "Enthusiastically ask the user if they want to build a Proposal Deal Deck or a custom Pitch Deck.",
                state, history
            )
            history.append({"user": message, "bot": deck_type_msg})
            return {"answer": deck_type_msg, "session_id": session_id, "sources": []}

        # C. If we previously asked to choose between deck types, accept the user's choice
        if history:
            last_bot = history[-1].get("bot", "").lower()
            if "proposal deal deck" in last_bot and "pitch deck" in last_bot:
                if "pitch" in msg_lower:
                    state["workflow"] = "pitch_deck"
                    state["phase"] = "gathering"
                    
                    # 👇 UPDATE 2: Pass 'message' and remove duplicate history append
                    init_q = handle_pitch_deck_workflow(message, state, history)
                    init_q.setdefault("session_id", session_id)
                    init_q.setdefault("sources", [])
                    return init_q

    # ---------------------------------------------
    # 🧩 Active Pitch Deck workflow
    # ---------------------------------------------
    # If a pitch deck flow is active, forward the message directly into that handler
    if state.get("workflow") == "pitch_deck":
        result = handle_pitch_deck_workflow(message, state, history)
        result.setdefault("session_id", session_id)
        result.setdefault("sources", [])
        return result

    # ---------------------------------------------
    # 💬 Smalltalk handler
    # ---------------------------------------------
    # Lightweight chit-chat (hello/thanks/etc.) gets a brief LLM response
    if is_smalltalk(message):
        try:
            reply = _safe_invoke_llm(
                _chat_once_is_smalltalk_prompt(message=message), 
                history=history, 
                compressed_context=state.get("compressed_context", "")
            ).strip()
        except Exception as e:
            logger.error("Smalltalk handling failed: %s", e)
            reply = generate_conversational_reply(
                "Playfully respond to their smalltalk, but remind them you are primarily here to build Pitch Decks and ask if they'd like to start one.",
                state, history
            )

        history.append({"user": message, "bot": reply})
        return {
        "answer": reply,
        "session_id": session_id,
        "sources": []
    }


    # ---------------------------------------------
    #  Out-of-scope handler (BLOCK BEFORE RAG)
    # ---------------------------------------------
    # Politely deflect questions that are unrelated to presentations and redirect the user
    if is_out_of_scope(message):
        comp_ctx = state.get("compressed_context", "")
        reply = _safe_invoke_llm(_chat_once_is_out_of_scope_prompt(message=message), history=history, compressed_context=comp_ctx).strip()

        history.append({"user": message, "bot": reply})
        return {
            "answer": reply,
            "session_id": session_id,
            "sources": []
        }

    # ---------------------------------------------
    # 🔍 RAG fallback (general Q&A)
    # ---------------------------------------------
    # If none of the above paths matched, try to answer via Retrieval-Augmented Generation
    logger.info("RAG fallback triggered for session %s — Query: %s", session_id, message)
    
    answer = run_qa(message, chat_history=chat_data)
    
    # If the RAG chain fails to find an answer, generate a natural fallback message
    if not answer or "don't know" in answer.lower() or "couldn't find" in answer.lower() or "not enough context" in answer.lower():
        fallback_situation = (
            "The user asked a question that was not found in your SharePoint documents. "
            "First, carefully check the conversation history to see if you can answer their question based on recent context "
            "(e.g., if they ask how many slides they requested, or what you just generated). "
            "If the answer is in the history, answer them directly and warmly. "
            "If the question is unrelated to building decks, about casual knowledge, trivia, jokes, or completely unknown, "
            "you MUST explicitly inform the user: 'I only help for deck generation about Intelliswift.' "
            "You may rephrase it slightly to sound natural, but you MUST convey that exact boundary without answering their question."
        )
        answer = generate_conversational_reply(fallback_situation, state, history)
        is_rag_fallback = True
    else:
        is_rag_fallback = False

    history.append({"user": message, "bot": answer})

    if is_new_session and not state.get("workflow"):
        if answer and not is_rag_fallback:
            return {"answer": answer, "session_id": session_id, "sources": []}
        greeting = generate_conversational_reply(
            "Give a warm, friendly greeting to the user. Tell them your name is Gia, Intelliswift's presentation assistant. Ask them if they'd like help building a pitch deck or have questions.",
            state, history
        )
        history.append({"user": message, "bot": greeting})
        return {"answer": greeting, "session_id": session_id, "sources": []}

    # Default return for RAG path when session/workflow already exists
    return {"answer": answer, "session_id": session_id, "sources": []}