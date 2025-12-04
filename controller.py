# controller.py
import os
import re
import json
from pathlib import Path
from typing import Optional, Tuple, List

# --- Robust intra-package imports (works as module and as script) ---
try:
    # When run as: python -m outlook_assist_V1.main
    from .model import (
        Language, GreetingStyle, SignOffStyle,
        GREETING_CATALOG, SIGNOFF_CATALOG,
        DEFAULT_SYSTEM_PROMPT, ReplyFormatConfig, MODEL_NAME
    )
except ImportError:
    # When run directly: python outlook_assist_V1/main.py
    import sys
    sys.path.append(os.path.dirname(__file__))
    from model import (
        Language, GreetingStyle, SignOffStyle,
        GREETING_CATALOG, SIGNOFF_CATALOG,
        DEFAULT_SYSTEM_PROMPT, ReplyFormatConfig, MODEL_NAME
    )

# Optional .env support
try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None


class OpenAIClientWrapper:
    """
    Wrapper for OpenAI API client that handles environment loading and sending chat completion requests.
    """
    def __init__(self):
        self._client = None

    def _load_env_and_client(self):
        """Load the OpenAI API key from environment or .env file and initialize the API client."""
        if load_dotenv is None:
            raise RuntimeError("Missing dependency 'python-dotenv'. Install with: pip install python-dotenv")
        try:
            from openai import OpenAI as OpenAIClient
        except Exception as e:
            raise RuntimeError("Missing or incompatible 'openai' package. Install openai>=1.0.0") from e

        env_key = os.getenv("OPENAI_API_KEY")
        if env_key:
            self._client = OpenAIClient(api_key=env_key)
            return

        # Search for .env file in common locations
        candidate_paths: List[Path] = []
        try:
            script_dir = Path(__file__).resolve().parent
            candidate_paths.append(script_dir / ".env")
        except Exception:
            pass
        candidate_paths.append(Path.cwd() / ".env")
        candidate_paths.append(Path.home() / ".env")

        for p in candidate_paths:
            try:
                if p.exists():
                    load_dotenv(dotenv_path=str(p))
                    env_key = os.getenv("OPENAI_API_KEY")
                    if env_key:
                        self._client = OpenAIClient(api_key=env_key)
                        return
            except Exception:
                continue

        searched = "\n".join([f"- {p} (exists: {p.exists()})" for p in candidate_paths])
        raise RuntimeError(
            "OPENAI_API_KEY not found in environment. Searched locations:\n"
            f"{searched}\n\n"
            "Fixes: Set OPENAI_API_KEY in a .env file or environment variable, and ensure 'python-dotenv' is installed."
        )

    def chat(self, system_prompt: str, user_prompt: str, model_name: str = MODEL_NAME) -> str:
        """Send a chat completion request and return the assistant's response content."""
        if self._client is None:
            self._load_env_and_client()
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
        try:
            resp = self._client.chat.completions.create(
                model=model_name, messages=messages, temperature=0.0
            )
        except Exception as e:
            raise RuntimeError("OpenAI API call failed: " + str(e))
        # Extract assistant content from response
        try:
            content = resp.choices[0].message.content
        except Exception:
            try:
                content = resp["choices"][0]["message"]["content"]
            except Exception:
                try:
                    content = resp.choices[0].message["content"]
                except Exception:
                    content = str(resp)
        if content is None:
            raise RuntimeError("Failed to retrieve content from OpenAI response.")
        return content


# single shared client
_openai_client = OpenAIClientWrapper()

# --------------------------
# Helpers & anti-parroting
# --------------------------

_WORD_RE = re.compile(r"[A-Za-zÆØÅæøåÉéÓóÚúÄäÖöÜüß0-9]+")

def _tokens(text: str) -> List[str]:
    return _WORD_RE.findall((text or "").lower())

def _ngram_set(tokens: List[str], n: int = 3) -> set:
    if n <= 0 or len(tokens) < n:
        return set()
    return {" ".join(tokens[i:i+n]) for i in range(len(tokens) - n + 1)}

def _parroting_ratio(src: str, out: str, n: int = 3) -> float:
    """
    Rough similarity test: n-gram overlap ratio.
    Returns 0..1 (higher means more likely the reply copied/transformed the source).
    """
    ts, to = _tokens(src), _tokens(out)
    As, Ao = _ngram_set(ts, n), _ngram_set(to, n)
    if not As or not Ao:
        return 0.0
    inter = len(As & Ao)
    denom = min(len(As), len(Ao))
    return inter / denom if denom else 0.0

def _looks_like_parroting(src: str, out: str) -> bool:
    """
    Heuristic:
      - 3-gram overlap > 0.35 (quite close),
      - and at least 10 overlapping 3-grams,
      - or long direct copy of 10+ consecutive tokens.
    """
    ratio = _parroting_ratio(src, out, n=3)
    ts, to = _tokens(src), _tokens(out)
    overlap_count = len(_ngram_set(ts, 3) & _ngram_set(to, 3))
    if ratio > 0.35 and overlap_count >= 10:
        return True
    # hard check: 10+ identical consecutive tokens
    joined_src = " ".join(ts)
    m = re.search(r"(\b\w+(?:\s+\w+){9,}\b)", " ".join(to))
    if m and m.group(0) in joined_src:
        return True
    return False

# --------------------------
# Empathic mirroring helpers
# --------------------------

_CASE_ID_RE = re.compile(r"(?:sag|case|reklamationssag)[^\d]{0,10}(\d{5,})", re.IGNORECASE)
_DAY_WINDOW_RE = re.compile(r"\b(\d{1,3})\s*(?:dage|days?)\b", re.IGNORECASE)

def _extract_case_ids(text: str) -> List[str]:
    if not text:
        return []
    ids = set(m.group(1) for m in _CASE_ID_RE.finditer(text))
    # Also capture long bare numbers that look like case ids
    ids |= set(m.group(0) for m in re.finditer(r"\b\d{6,}\b", text))
    return list(ids)

def _extract_day_windows(text: str) -> List[str]:
    if not text:
        return []
    found = [m.group(1) for m in _DAY_WINDOW_RE.finditer(text)]
    norm = []
    for d in found:
        try:
            n = int(d)
            if 1 <= n <= 365:
                norm.append(f"{n} days")
        except Exception:
            pass
    seen = set()
    out = []
    for x in norm:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out[:3]

def _build_mirroring_hint(incoming_body: str) -> str:
    """
    Build a short list of facts the model should mirror back empathetically in its own words.
    Keep this list tiny so it doesn't encourage parroting.
    """
    facts = []
    case_ids = _extract_case_ids(incoming_body or "")
    if case_ids:
        facts.append(f"case number: {case_ids[0]}")
    days = _extract_day_windows(incoming_body or "")
    if days:
        if len(days) == 1:
            facts.append(f"estimated timeline: ~{days[0]}")
        elif len(days) >= 2:
            facts.append(f"estimated timeline: ~{days[0]} to {days[1]}")
    if re.search(r"opdatering|update|via e-?mail|e-mail", incoming_body or "", re.IGNORECASE):
        facts.append("you’ll send updates via email")

    if not facts:
        return ""
    return "Mirror (in your own words, no quotes) these facts in a single short sentence: " + "; ".join(facts) + "."

# --------------------------
# Extraction & formatting
# --------------------------

def extract_sender_name_from_signature(body_text: str) -> Optional[str]:
    """Attempt to extract the sender's name from the email signature in the given body text."""
    if not body_text:
        return None
    lines = [ln.strip() for ln in body_text.strip().splitlines() if ln.strip()]
    last_lines = lines[-12:] if len(lines) > 12 else lines
    block = "\n".join(last_lines)
    markers = [
        r"Med venlig hilsen[,:\-]?\s*\n?(?P<name>.+)",
        r"Mvh[,:\-]?\s*\n?(?P<name>.+)",
        r"Venlig hilsen[,:\-]?\s*\n?(?P<name>.+)",
        r"Best regards[,:\-]?\s*\n?(?P<name>.+)",
        r"Regards[,:\-]?\s*\n?(?P<name>.+)",
        r"Kind regards[,:\-]?\s*\n?(?P<name>.+)"
    ]
    for pat in markers:
        m = re.search(pat, block, flags=re.IGNORECASE | re.DOTALL)
        if m:
            name = m.group("name")
            if name:
                name_line = name.strip().splitlines()[0].strip()
                name_line = re.sub(r"<.*?>", "", name_line)
                name_line = re.sub(r"[^A-Za-zÆØÅæøåéÉóÓúÚäÄöÖüÜß \-\.'\"]", "", name_line).strip()
                if name_line:
                    return name_line
    if last_lines:
        cand = last_lines[-1]
        if len(cand.split()) <= 4 and "@" not in cand and not re.search(r"\d{2,}", cand):
            return re.sub(r"[^A-Za-zÆØÅæøåéÉóÓúÚäÄöÖüÜß \-\.'\"]", "", cand).strip()
    return None

def build_reply_prompt_json(recipient_name: str,
                            sender_name_hint: Optional[str],
                            incoming_subject: Optional[str],
                            incoming_body: str,
                            tone_instructions: Optional[str],
                            extra_instructions: Optional[str],
                            mirroring_hint: str,
                            reply_language: Language) -> str:
    """Construct the user prompt for the model to generate a JSON-formatted reply."""
    sender_ref = sender_name_hint if sender_name_hint else "(the sender)"
    tone = tone_instructions or "concise and professional"

    # Language-specific instructions/examples
    lang_name = "DANISH" if reply_language == Language.DA else "ENGLISH"
    # Example empathy micro-templates by language (only as guidance to the model)
    empathy_example = (
        "I appreciate the update on case 123456; I’ve noted the expected 14–30 day timeline."
        if reply_language != Language.DA else
        "Tak for opdateringen vedrørende sag 123456; jeg har noteret den forventede behandlingstid på ca. 14–30 dage."
    )

    # Strong anti-parroting + reply outline + empathy
    prompt_lines = [
        f"ACT AS: You are writing an email REPLY as the recipient: {recipient_name or '[recipient name]'}.\n",
        "CRITICAL RULES:",
        "1) DO NOT restate, translate, or quote the original message. Avoid copying phrases from it.",
        "2) Write ONLY from the recipient's perspective (use 'I' or 'we' / 'jeg' or 'vi').",
        "3) NEVER sign or speak as the sender or the sender's organization.",
        f"4) Reply in {lang_name}.",
        "5) Output **only** a single JSON object. No extra text.",
        "",
        "EMPATHY & MIRRORING:",
        "- Include ONE short, natural line that reflects the most important detail(s) you noticed (e.g., case number or timeline) in your own words.",
        "- This mirroring line must be 1 sentence, and must NOT quote or translate exact phrases.",
        f"- Example (do NOT copy; adapt to the context & language): \"{empathy_example}\"",
        "",
        "REPLY OUTLINE (guidance, adapt as needed):",
        "- Subject: 'Re: [incoming subject]' (or a clear variant).",
        "- Greeting to the sender by name if available.",
        "- Empathic acknowledgement line (1 sentence) using the most important details.",
        "- Next step(s) or what you will do / need from them (confirm info, ask a clarifying question, set expectations).",
        "- Close politely.",
        "",
        "OUTPUT SCHEMA:",
        '{"subject":"...", "body":"..."}',
        "",
        "CONTEXT (incoming email):",
        f"Sender (detected/provided): {sender_ref}",
        f"Subject: {incoming_subject or '[no subject]'}",
        "Body:",
        incoming_body or "[no body]",
        "",
        "REPLY REQUIREMENTS:",
        f"- Tone/Style: {tone}.",
        f"- Sign-off example: {'Best regards,' if reply_language != Language.DA else 'Med venlig hilsen,'} {recipient_name or '[Your Name]'}",
    ]
    if extra_instructions:
        prompt_lines += ["", "ADDITIONAL NOTES:", extra_instructions]
    if mirroring_hint:
        prompt_lines += ["", "SALIENT FACTS (for the one-sentence mirroring):", mirroring_hint]

    prompt_lines += ["", "Now produce the reply as a JSON object only."]
    return "\n".join(prompt_lines)

def _extract_json_object(text: str) -> Optional[dict]:
    """Extract the largest JSON object from text (ignoring Markdown formatting)."""
    if not text:
        return None
    try:
        obj = json.loads(text.strip())
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    clean = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.IGNORECASE)
    best_obj = None
    best_len = 0
    for m in re.finditer(r"\{", clean):
        start = m.start()
        depth = 0
        for i in range(start, len(clean)):
            if clean[i] == "{":
                depth += 1
            elif clean[i] == "}":
                depth -= 1
                if depth == 0:
                    candidate = clean[start:i+1]
                    try:
                        obj = json.loads(candidate)
                        if isinstance(obj, dict) and len(candidate) > best_len:
                            best_obj = obj
                            best_len = len(candidate)
                    except json.JSONDecodeError:
                        pass
                    break
    return best_obj

def parse_model_json_output(output_text: str) -> Tuple[Optional[str], Optional[str]]:
    """Parse the model's output for a JSON object and return its 'subject' and 'body'."""
    data = _extract_json_object(output_text)
    if not data:
        return None, None
    subj = (data.get("subject") or data.get("Subject") or "").strip()
    body = (data.get("body") or data.get("Body") or "").strip()
    return (subj if subj else None, body if body else None)

def _normalize_name(name: str) -> str:
    """Normalize a name by collapsing multiple whitespace characters."""
    return re.sub(r"\s+", " ", (name or "").strip())

def choose_greeting(cfg: ReplyFormatConfig, sender_name: Optional[str]) -> str:
    """Choose an appropriate greeting based on configuration and sender_name."""
    style = cfg.greeting_style
    if style == GreetingStyle.AUTO:
        style = GreetingStyle.FORMAL if sender_name else GreetingStyle.NEUTRAL
    template = GREETING_CATALOG.get(cfg.language, {}).get(style)
    if not template:
        template = GREETING_CATALOG[Language.EN][GreetingStyle.NEUTRAL]
    safe_name = _normalize_name(sender_name or "")
    try:
        return template.format(name=safe_name) if "{name}" in template else template
    except Exception:
        return template

def choose_signoff(cfg: ReplyFormatConfig) -> str:
    """Choose a sign-off phrase based on configuration."""
    phrase = SIGNOFF_CATALOG.get(cfg.language, {}).get(cfg.signoff_style)
    if not phrase:
        phrase = SIGNOFF_CATALOG[Language.EN][SignOffStyle.BEST_REGARDS]
    return phrase

def has_initial_greeting(text: str) -> bool:
    """Check if text already starts with a greeting phrase."""
    if not text:
        return False
    lines = [ln.strip().lower() for ln in text.lstrip().splitlines() if ln.strip()]
    return bool(lines) and lines[0].startswith(("dear ", "hello", "hi ", "hej", "kære"))

def has_signoff(text: str, recipient_display_name: str) -> bool:
    """Check if text already contains a sign-off with the recipient's name."""
    if not text:
        return False
    lines = [ln.strip().lower() for ln in text.strip().splitlines() if ln.strip()]
    if len(lines) < 2:
        return False
    last, second_last = lines[-1], lines[-2]
    recip = (recipient_display_name or "").strip().lower()
    return bool(recip and recip in last and any(k in second_last for k in ("regards", "hilsen", "cheers", "thanks", "venlig")))

def inject_greeting_and_signoff(body: str, cfg: ReplyFormatConfig, sender_name: Optional[str], recipient_display_name: str) -> str:
    """Ensure the email body has a proper greeting at start and sign-off at end."""
    text = body or ""
    if not has_initial_greeting(text):
        greet_line = choose_greeting(cfg, sender_name)
        text = f"{greet_line}\n" + ("\n" * max(1, cfg.blank_lines_after_greeting)) + text.lstrip()
    if not has_signoff(text, recipient_display_name):
        signoff_line = choose_signoff(cfg)
        text = text.rstrip() + ("\n" * max(1, cfg.blank_lines_before_signoff))
        text += f"{signoff_line}\n{recipient_display_name}"
    return text

def _signs_off_with_sender(body_text: str, sender_name: str) -> bool:
    """Detect if the draft reply erroneously signs off with the sender's name."""
    if not body_text or not sender_name:
        return False
    sender_low = sender_name.strip().lower()
    lines = [ln.strip().lower() for ln in body_text.splitlines() if ln.strip()]
    if not lines:
        return False
    last_line = lines[-1]
    if sender_low and (last_line == sender_low or last_line.startswith(sender_low.split()[0])):
        return True
    if len(lines) >= 2:
        second_last = lines[-2]
        if sender_low in last_line and any(k in second_last for k in ("regards", "hilsen", "cheers", "thanks", "venlig")):
            return True
    return False

def generate_reply_email(recipient_display_name: str,
                         incoming_sender_name: str,
                         incoming_sender_email: str,
                         incoming_subject: str,
                         incoming_body: str,
                         tone_instructions: str = "",
                         extra_instructions: str = "",
                         greeting_style: GreetingStyle = GreetingStyle.AUTO,
                         signoff_style: SignOffStyle = SignOffStyle.BEST_REGARDS,
                         language: Language = Language.EN) -> Tuple[str, str]:
    """Generate the reply email subject and body content for a given incoming email."""
    # Determine sender name (use provided or extract from body)
    sender_name = incoming_sender_name or extract_sender_name_from_signature(incoming_body)

    # Build a tiny mirroring hint from the incoming email
    mirroring_hint = _build_mirroring_hint(incoming_body or "")

    # Build prompts for OpenAI (language-aware)
    user_prompt = build_reply_prompt_json(
        recipient_name=recipient_display_name,
        sender_name_hint=sender_name,
        incoming_subject=incoming_subject,
        incoming_body=incoming_body,
        tone_instructions=tone_instructions.strip(),
        extra_instructions=extra_instructions.strip(),
        mirroring_hint=mirroring_hint,
        reply_language=language
    )

    system_prompt = (
        f"ROLE: You must reply AS the recipient named {recipient_display_name}. "
        "MANDATES: Do not restate/translate the original; do not speak as the sender; first-person only. "
        "OUTPUT: JSON only as per the user prompt."
    )

    # Get model output
    output = _openai_client.chat(system_prompt, user_prompt, model_name=MODEL_NAME)
    subj, body = parse_model_json_output(output)
    if not subj and not body:
        # Attempt manual parsing if JSON wasn't returned properly
        low = output.lower()
        idx_sub = low.find("subject:")
        idx_bod = low.find("\nbody:")
        if idx_sub != -1 and idx_bod != -1 and idx_sub < idx_bod:
            subj = output[idx_sub + len("subject:"): idx_bod].strip()
            body = output[idx_bod + len("\nbody:"):].strip()
        else:
            subj = ("Re: " + (incoming_subject or "")).strip()
            body = output.strip()

    # -------------------------
    # Anti-parroting safeguard
    # -------------------------
    if body and _looks_like_parroting(incoming_body or "", body):
        correction_prompt = (
            "Your previous reply copied/transformed the original message (parroting). "
            "Regenerate the JSON reply with these constraints:\n"
            " - DO NOT copy or translate phrases from the original email.\n"
            " - Include ONE short empathy line that mirrors key facts in your own words.\n"
            f" - Sign as the recipient: {recipient_display_name}.\n"
            " - Keep it concise and professional.\n\n"
            "Reuse the same output schema.\n\n"
            "Original task and context below:\n\n" + user_prompt
        )
        try:
            retry_output = _openai_client.chat(system_prompt, correction_prompt, model_name=MODEL_NAME)
            new_subj, new_body = parse_model_json_output(retry_output)
            if new_body:  # accept only if we actually get a new body
                subj = new_subj or subj
                body = new_body
        except Exception:
            pass

    # Also fix if it signs off with sender's name
    if sender_name and body and _signs_off_with_sender(body, sender_name):
        correction_prompt = (
            "The previous JSON reply mistakenly signed as the SENDER. "
            f"Please regenerate the reply JSON and sign as the RECIPIENT ({recipient_display_name}). "
            "Write the reply from the recipient's perspective. Do not copy or translate the original. JSON only.\n\n"
            "Original prompt and email context:\n\n" + user_prompt
        )
        try:
            retry_output = _openai_client.chat(system_prompt, correction_prompt, model_name=MODEL_NAME)
            new_subj, new_body = parse_model_json_output(retry_output)
            if new_subj or new_body:
                subj = new_subj or subj
                body = new_body or body
        except Exception:
            pass

    # Ensure greeting and sign-off are in place (language-aware)
    cfg = ReplyFormatConfig(
        language=language,
        greeting_style=greeting_style,
        signoff_style=signoff_style,
        blank_lines_after_greeting=1,
        blank_lines_before_signoff=2
    )
    body = inject_greeting_and_signoff(body or "", cfg, sender_name, recipient_display_name)
    final_subj = (subj or ("Re: " + (incoming_subject or ""))).strip()
    return final_subj, body

def generate_new_email(recipient_address: str, subject_line: str, topic_description: str) -> str:
    """Generate a new email (body text) for the given topic description."""
    if not topic_description:
        raise ValueError("Topic description cannot be empty.")
    prompt = (
        f"Draft a professional email to {recipient_address or '[recipient]'} about:\n\n{topic_description}\n\n"
        "Include a clear subject line suggestion and a polite sign-off."
    )
    return _openai_client.chat(DEFAULT_SYSTEM_PROMPT, prompt, model_name=MODEL_NAME)
