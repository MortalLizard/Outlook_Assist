from enum import Enum
from dataclasses import dataclass
from typing import Dict

# Model layer: contains static data definitions (enums, catalogs, configuration dataclasses)

MODEL_NAME = "gpt-4"

class Language(Enum):
    EN = "en"
    DA = "da"

class GreetingStyle(Enum):
    AUTO = "auto"
    FORMAL = "formal"
    NEUTRAL = "neutral"
    CASUAL = "casual"

class SignOffStyle(Enum):
    BEST_REGARDS = "best_regards"
    KIND_REGARDS = "kind_regards"
    REGARDS = "regards"
    CHEERS = "cheers"
    THANKS = "thanks"

# Catalogs for greeting and sign-off phrases by language and style
GREETING_CATALOG: Dict[Language, Dict[GreetingStyle, str]] = {
    Language.EN: {
        GreetingStyle.FORMAL: "Dear {name},",
        GreetingStyle.NEUTRAL: "Hello,",
        GreetingStyle.CASUAL: "Hi {name},",
    },
    Language.DA: {
        GreetingStyle.FORMAL: "Kære {name},",
        GreetingStyle.NEUTRAL: "Hej,",
        GreetingStyle.CASUAL: "Hej {name},",
    },
}

SIGNOFF_CATALOG: Dict[Language, Dict[SignOffStyle, str]] = {
    Language.EN: {
        SignOffStyle.BEST_REGARDS: "Best regards,",
        SignOffStyle.KIND_REGARDS: "Kind regards,",
        SignOffStyle.REGARDS: "Regards,",
        SignOffStyle.CHEERS: "Cheers,",
        SignOffStyle.THANKS: "Thanks,",
    },
    Language.DA: {
        SignOffStyle.BEST_REGARDS: "Med venlig hilsen,",
        SignOffStyle.KIND_REGARDS: "Venlig hilsen,",
        SignOffStyle.REGARDS: "Hilsen,",
        SignOffStyle.CHEERS: "De bedste hilsner,",
        SignOffStyle.THANKS: "Tak,",
    },
}

# Default system prompt for the OpenAI assistant
# STRONG anti-parroting + POV constraints — used for generic drafting (non-reply path).
DEFAULT_SYSTEM_PROMPT = (
    "ROLE: You write email DRAFTS strictly as the RECIPIENT (first person 'I' or 'we'). "
    "MANDATES:\n"
    " - NEVER restate, translate, or summarize the original message; compose a response that advances the thread.\n"
    " - NEVER write as the sender; NEVER sign or speak as the sender's organization.\n"
    " - Keep it concise, professional, and actionable.\n"
    " - Prefer ENGLISH unless otherwise explicitly requested.\n"
    "OUTPUT: Provide only the text requested by the user/invoking prompt."
)

@dataclass
class ReplyFormatConfig:
    language: Language = Language.EN
    greeting_style: GreetingStyle = GreetingStyle.AUTO
    signoff_style: SignOffStyle = SignOffStyle.BEST_REGARDS
    blank_lines_after_greeting: int = 1
    blank_lines_before_signoff: int = 1
