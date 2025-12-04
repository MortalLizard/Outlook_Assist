# Outlook Assist - AI-Powered Email Response Generator

## Overview

Outlook Assist is an intelligent email assistant that reads incoming emails and automatically generates professional, contextually appropriate response drafts. The system uses OpenAI's GPT-4 to analyze email content and compose empathetic, personalized replies while avoiding common pitfalls like parroting or incorrect attribution.

## Key Features

- **Automatic Reply Generation**: Analyzes incoming emails and generates appropriate responses
- **Multi-language Support**: Supports English (EN) and Danish (DA)
- **Smart Context Understanding**: Extracts key information like case numbers, timelines, and sender names
- **Anti-Parroting Protection**: Built-in safeguards prevent simply restating the original email
- **Customizable Tone**: Configure greeting styles, sign-offs, and tone
- **Multiple Interfaces**: Command-line tool, Flask API server, and Outlook add-in
- **Outlook Integration**: Direct integration with Microsoft Outlook on Windows

---

## Project Architecture

The project consists of several interconnected components:

### Core Python Modules

#### 1. **model.py** - Data Models and Configuration
Defines the core data structures and catalogs:
- **Language**: Enum for EN (English) and DA (Danish)
- **GreetingStyle**: FORMAL, NEUTRAL, CASUAL, AUTO
- **SignOffStyle**: BEST_REGARDS, KIND_REGARDS, REGARDS, CHEERS, THANKS
- **Greeting/SignOff Catalogs**: Multilingual phrase templates
- **ReplyFormatConfig**: Dataclass for reply formatting preferences
- **Default system prompts**: Instructions for the AI model

#### 2. **controller.py** - Core Business Logic
The brain of the application containing:

**OpenAI Client Wrapper**:
- Manages OpenAI API authentication (loads from `.env` file)
- Sends chat completion requests to GPT-4
- Handles errors and response parsing

**Key Functions**:
- `generate_reply_email()`: Main function to generate email replies
- `extract_sender_name_from_signature()`: Extracts sender name from email signatures
- `build_reply_prompt_json()`: Constructs detailed prompts for the AI model
- `parse_model_json_output()`: Parses JSON responses from GPT-4
- `inject_greeting_and_signoff()`: Ensures proper greeting/closing in replies
- `generate_new_email()`: Creates brand new emails (not replies)

**Anti-Parroting System**:
- `_parroting_ratio()`: Calculates n-gram overlap between source and generated text
- `_looks_like_parroting()`: Heuristic detection of copied content
- Automatic regeneration if parroting is detected

**Empathic Mirroring**:
- `_extract_case_ids()`: Finds case/reference numbers in emails
- `_extract_day_windows()`: Detects timelines (e.g., "14 days")
- `_build_mirroring_hint()`: Creates prompts for acknowledging key facts naturally

#### 3. **main.py** - Command-Line Interface
Interactive CLI tool for email composition:
- **Reply Mode** (`--reply`): Generate replies to incoming emails
- **New Email Mode**: Draft new emails from scratch
- Multi-line input support (end with `.`)
- Outlook integration via COM (opens draft in Outlook)
- Configurable greeting/sign-off styles and language

#### 4. **outlook_utils.py** - Windows Outlook Integration
Provides COM automation for Microsoft Outlook:
- `create_outlook_email()`: Opens a new Outlook draft window
- Uses `win32com.client` (pywin32) for Windows COM interface
- Automatically populates To, Subject, and Body fields

#### 5. **server.py** - Flask API Server
HTTP API for integration with web-based add-ins:
- **POST `/assist/reply`**: Generate email replies (main endpoint)
- **GET `/healthz`**: Health check endpoint
- Accepts JSON payloads with email context
- Returns JSON with subject and body

### Web Components (Outlook Add-in)

#### 6. **taskpane.ts** - TypeScript Client
Office.js-based Outlook add-in logic:
- `OutlookAiClient`: Communicates with Flask server
- `OutlookAddinApp`: Manages add-in UI and Outlook integration
- Reads current email item (sender, subject, body)
- Calls API to generate draft
- Injects draft into compose window or opens reply form

#### 7. **manifest.xml** - Outlook Add-in Manifest
Defines the Outlook add-in configuration:
- Extension points for Read and Compose modes
- Ribbon buttons for "Draft with AI"
- Permissions and runtime requirements
- Points to taskpane.html and commands.js

#### 8. **tsconfig.json** - TypeScript Configuration
Compiler settings for TypeScript files:
- Target: ES2019
- Module resolution: Bundler
- Types: office-js
- Strict mode enabled

---

## How It Works: Email Reply Flow

### Step-by-Step Process

#### 1. **Input Collection**
The system gathers information about the incoming email:
- Recipient's display name (who is replying)
- Sender's name and email address
- Email subject line
- Email body content
- Optional tone/style instructions
- Language and formatting preferences

#### 2. **Context Analysis** (controller.py)
The system analyzes the incoming email:

**Sender Identification**:
```python
extract_sender_name_from_signature(body_text)
```
- Scans the email signature for sender name
- Uses regex patterns for common sign-offs ("Best regards", "Med venlig hilsen")
- Falls back to heuristics if no clear signature found

**Fact Extraction**:
```python
_build_mirroring_hint(incoming_body)
```
- Extracts case/reference numbers (e.g., "Case 123456")
- Identifies timelines (e.g., "14-30 days")
- Notes communication preferences (e.g., "via email")
- Creates concise hints for empathetic acknowledgment

#### 3. **Prompt Construction**
The system builds a detailed prompt for GPT-4:

```python
build_reply_prompt_json(
    recipient_name="John Doe",
    sender_name_hint="Jane Smith",
    incoming_subject="Order Update",
    incoming_body="[email content]",
    tone_instructions="professional and concise",
    mirroring_hint="case number: 123456; timeline: ~14 days",
    reply_language=Language.EN
)
```

**The prompt includes**:
- Clear role definition (replying AS the recipient)
- Critical rules (no parroting, no quoting, first-person only)
- Empathy guidelines with examples
- Output format specification (JSON: {subject, body})
- Full context of incoming email
- Language-specific instructions

#### 4. **AI Generation**
```python
_openai_client.chat(system_prompt, user_prompt, model="gpt-4")
```
- Sends the prompt to OpenAI GPT-4
- Temperature set to 0.0 for consistent output
- Expects JSON response: `{"subject": "...", "body": "..."}`

#### 5. **Quality Checks**

**Anti-Parroting Detection**:
```python
if _looks_like_parroting(incoming_body, generated_body):
    # Regenerate with stronger constraints
```
- Calculates 3-gram overlap ratio
- Checks for long sequences of copied tokens
- Triggers regeneration if threshold exceeded (>35% overlap + 10+ n-grams)

**Signature Verification**:
```python
if _signs_off_with_sender(body, sender_name):
    # Regenerate to fix incorrect attribution
```
- Ensures reply is signed by the recipient, not the sender
- Prevents role confusion

#### 6. **Formatting**
```python
inject_greeting_and_signoff(body, config, sender_name, recipient_name)
```
- Adds appropriate greeting if missing ("Dear Jane," / "Hej Jane,")
- Ensures proper sign-off with recipient's name
- Respects language and style preferences
- Adds appropriate blank lines for readability

#### 7. **Output Delivery**
The formatted reply is returned/displayed:
- **CLI**: Prints to console and opens in Outlook
- **API**: Returns JSON to client
- **Add-in**: Injects into compose window or opens reply form

---

## Usage Examples

### Command-Line Interface

#### Generate a Reply
```bash
python main.py --reply --lang en --greeting formal --signoff best_regards
```

**Interactive prompts**:
1. Your name: `John Doe`
2. Sender's name: `Jane Smith`
3. Sender's email: `jane@example.com`
4. Subject: `Order Status Update`
5. Email body: `[paste email content, end with .]`
6. Tone instructions: `professional and concise`
7. Extra notes: `[optional, end with .]`

**Output**:
- Displays generated reply
- Opens draft in Outlook (if available)
- Option to include original message

#### Draft New Email
```bash
python main.py
```
- Prompts for recipient, subject, and topic description
- Generates email from scratch
- Opens in Outlook

### Flask API Server

#### Start the Server
```bash
python server.py
```
Server runs on `http://127.0.0.1:5001`

#### API Request Example
```bash
curl -X POST http://127.0.0.1:5001/assist/reply \
  -H "Content-Type: application/json" \
  -d '{
    "recipient_display_name": "John Doe",
    "incoming_sender_name": "Jane Smith",
    "incoming_sender_email": "jane@example.com",
    "incoming_subject": "Order Update",
    "incoming_body": "Your order #123456 will be processed in 14-30 days.",
    "tone": "professional",
    "greeting_style": "auto",
    "signoff_style": "best_regards",
    "language": "en"
  }'
```

**Response**:
```json
{
  "subject": "Re: Order Update",
  "body": "Dear Jane,\n\nThank you for the update on order #123456..."
}
```

### Outlook Add-in

1. Install the add-in in Outlook (sideload manifest.xml)
2. Open an email in Outlook
3. Click "Draft with AI" button in the ribbon
4. System reads the email, generates reply, and opens compose window

---

## Configuration

### Environment Setup

Create a `.env` file in the project root:
```env
OPENAI_API_KEY=sk-your-api-key-here
```

The system searches for `.env` in:
1. Script directory
2. Current working directory
3. User's home directory

### Reply Format Configuration

Customize in code using `ReplyFormatConfig`:
```python
cfg = ReplyFormatConfig(
    language=Language.DA,              # or Language.EN
    greeting_style=GreetingStyle.FORMAL,
    signoff_style=SignOffStyle.BEST_REGARDS,
    blank_lines_after_greeting=1,
    blank_lines_before_signoff=2
)
```

---

## Installation

### Prerequisites
- Python 3.8+
- OpenAI API key
- (Optional) Microsoft Outlook for Windows
- (Optional) Node.js/npm for TypeScript compilation

### Python Dependencies
```bash
pip install openai python-dotenv flask pywin32
```

### TypeScript Dependencies (for add-in)
```bash
npm install --save-dev typescript @types/office-js
```

### Compile TypeScript
```bash
npx tsc
```

---

## Technical Highlights

### 1. **Robust Anti-Parroting System**
The system uses n-gram analysis to detect when the AI is simply rephrasing the original email. If detected, it automatically regenerates with stronger constraints.

### 2. **Empathic Mirroring**
Rather than ignoring the original email, the system extracts key facts (case numbers, timelines) and prompts the AI to acknowledge them naturally in its own words.

### 3. **Language-Aware Formatting**
All greeting/sign-off templates are localized for both English and Danish, with automatic selection based on context.

### 4. **Multi-Interface Architecture**
The same core logic (controller.py) powers three different interfaces:
- CLI for quick terminal usage
- Flask API for web integration
- Outlook add-in for seamless workflow

### 5. **Intelligent Name Extraction**
Uses multiple strategies to identify sender names from signatures, including regex patterns for common phrases and fallback heuristics.

### 6. **JSON-Structured Output**
The AI is instructed to return structured JSON, making parsing reliable and enabling easy integration with other systems.

---

## Project Structure

```
Outlook_Assist/
├── main.py              # CLI entry point
├── controller.py        # Core business logic
├── model.py            # Data models and enums
├── outlook_utils.py    # Outlook COM integration
├── server.py           # Flask API server
├── taskpane.ts         # Outlook add-in TypeScript
├── taskpane.html       # Add-in UI (referenced but empty)
├── commands.ts         # Add-in commands (empty)
├── manifest.xml        # Outlook add-in manifest
├── tsconfig.json       # TypeScript configuration
├── .env                # Environment variables (not in git)
├── .gitignore          # Git ignore patterns
└── README.md           # This file
```

---

## Error Handling

The system includes comprehensive error handling:
- **Missing API Key**: Clear error messages with search locations
- **API Failures**: Catches and reports OpenAI API errors
- **Outlook Not Available**: Graceful fallback with manual copy instructions
- **Invalid Input**: Validation for required fields
- **Parroting Detection**: Automatic retry with corrected prompts
- **JSON Parsing**: Fallback to manual parsing if JSON is malformed

---

## Development Notes

### Why GPT-4?
The project uses GPT-4 (specified in `model.py`) for superior reasoning about context, tone, and natural language generation compared to earlier models.

### COM Integration Limitations
`outlook_utils.py` uses Windows COM automation, which:
- Only works on Windows
- Requires Outlook to be installed
- Needs `pywin32` package

### TypeScript Type Safety
The add-in uses type casting (`as any`) in some places to work around Office.js type definition inconsistencies across versions.

---

## Future Enhancements

Potential improvements:
- Support for additional languages
- Template library for common response types
- Learning from user corrections
- Attachment handling
- Calendar integration for meeting scheduling
- Sentiment analysis for tone adjustment
- Multi-thread conversation context

---

## Security Considerations

- **API Key**: Store in `.env`, never commit to version control
- **Email Content**: Sent to OpenAI API (review OpenAI's data policies)
- **Local Processing**: Consider using local LLMs for sensitive data
- **HTTPS**: Use SSL for production Flask server
- **Input Validation**: Sanitize user inputs in production deployments

---

## License

[Specify your license here]

---

## Support

For issues or questions:
1. Check error messages for specific guidance
2. Verify `.env` configuration
3. Ensure all dependencies are installed
4. Check OpenAI API quota and status

---

## Acknowledgments

Built with:
- OpenAI GPT-4 API
- Flask web framework
- Office.js for Outlook integration
- pywin32 for Windows COM automation
