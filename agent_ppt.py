"""
agent_ppt.py
------------
The Auto-PPT Agent.

Architecture:
  User Prompt
      │
      ▼
  LangChain ReAct Agent (Google Gemini brain)
      │
      ├── MCP Tools from pptx_mcp_server.py   (imported directly)
      └── MCP Tools from search_mcp_server.py (imported directly)

Config: All settings loaded from .env file.
Usage:
  python agent_ppt.py "Create a 5-slide presentation on the life cycle of a star"
  python agent_ppt.py   (interactive mode)
"""

import json
import logging
import os
import sys
from typing import Any, Optional, List

# ── Load .env FIRST ───────────────────────────────────────────────────────────
from dotenv import load_dotenv
load_dotenv()

# ── Validate required env var ─────────────────────────────────────────────────
if not os.getenv("GOOGLE_API_KEY"):
    print("❌  GOOGLE_API_KEY is not set in your .env file!")
    print("    Get a free key at: aistudio.google.com/apikey")
    sys.exit(1)

# ── Read all config from .env (with defaults) ─────────────────────────────────
MODEL_ID          = os.getenv("MODEL_ID",             "gemini-2.0-flash")
MAX_TOKENS        = int(os.getenv("MAX_TOKENS",       "1024"))
TEMPERATURE       = float(os.getenv("TEMPERATURE",    "0.2"))
VECTOR_STORE_PATH = os.getenv("VECTOR_STORE_PATH",    "./memory")
K_RETRIEVAL       = int(os.getenv("K_RETRIEVAL",      "3"))
AGENT_MAX_ITER    = int(os.getenv("AGENT_MAX_ITERATIONS", "15"))
AGENT_VERBOSE     = os.getenv("AGENT_VERBOSE", "false").lower() == "true"
LOG_LEVEL         = os.getenv("LOG_LEVEL",            "INFO")
AGENT_WORKSPACE   = os.getenv("AGENT_WORKSPACE",      "./workspace")

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper(), logging.INFO),
    format="%(asctime)s [%(levelname)s] %(name)s – %(message)s",
)
logger = logging.getLogger("auto-ppt-agent")

# ── Ensure folders exist ──────────────────────────────────────────────────────
os.makedirs(AGENT_WORKSPACE, exist_ok=True)
os.makedirs(VECTOR_STORE_PATH, exist_ok=True)

logger.info("Config loaded ✅")
logger.info(f"  MODEL_ID          = {MODEL_ID}")
logger.info(f"  MAX_TOKENS        = {MAX_TOKENS}")
logger.info(f"  TEMPERATURE       = {TEMPERATURE}")
logger.info(f"  AGENT_MAX_ITER    = {AGENT_MAX_ITER}")
logger.info(f"  AGENT_VERBOSE     = {AGENT_VERBOSE}")
logger.info(f"  AGENT_WORKSPACE   = {AGENT_WORKSPACE}")
logger.info(f"  VECTOR_STORE_PATH = {VECTOR_STORE_PATH}")
logger.info(f"  K_RETRIEVAL       = {K_RETRIEVAL}")

# ── LangChain ─────────────────────────────────────────────────────────────────
from langchain.agents import AgentExecutor, create_react_agent
from langchain.tools import Tool
from langchain_core.prompts import PromptTemplate
from langchain_google_genai import ChatGoogleGenerativeAI


# ── Import MCP tool functions directly from server files ──────────────────────
logger.info("Loading MCP tools from pptx_mcp_server...")
from pptx_mcp_server import (
    create_presentation,
    add_slide,
    write_text,
    add_image_placeholder,
    save_presentation,
    get_slide_count,
)

logger.info("Loading MCP tools from search_mcp_server...")
from search_mcp_server import search_topic

logger.info("✅ All MCP tools loaded successfully!")


# ──────────────────────────────────────────────────────────────────────────────
# LANGCHAIN TOOL WRAPPERS
# ──────────────────────────────────────────────────────────────────────────────

def _parse_args(args_str: str) -> dict:
    """Safely parse agent's JSON string argument into a Python dict."""
    args_str = args_str.strip()
    if not args_str or args_str in ("{}", ""):
        return {}
    try:
        return json.loads(args_str)
    except json.JSONDecodeError:
        try:
            return json.loads(args_str.replace("'", '"'))
        except Exception:
            try:
                return eval(args_str)  # noqa: S307
            except Exception:
                return {}


def tool_create_presentation(args_str: str) -> str:
    args = _parse_args(args_str)
    output_path = args.get("output_path", f"{AGENT_WORKSPACE}/presentation.pptx")
    if not output_path.startswith(AGENT_WORKSPACE):
        output_path = f"{AGENT_WORKSPACE}/{os.path.basename(output_path)}"
    return create_presentation(output_path)


def tool_add_slide(args_str: str) -> str:
    args = _parse_args(args_str)
    return add_slide(
        slide_type = args.get("slide_type", "content"),
        title      = args.get("title", "Slide"),
        subtitle   = args.get("subtitle", ""),
        bullets    = args.get("bullets", []),
    )


def tool_write_text(args_str: str) -> str:
    args = _parse_args(args_str)
    return write_text(
        slide_index = int(args.get("slide_index", 0)),
        title       = args.get("title", ""),
        bullets     = args.get("bullets", []),
    )


def tool_add_image_placeholder(args_str: str) -> str:
    args = _parse_args(args_str)
    return add_image_placeholder(
        slide_index = int(args.get("slide_index", 0)),
        label       = args.get("label", "[Image]"),
        position    = args.get("position", "right"),
    )


def tool_save_presentation(args_str: str) -> str:
    return save_presentation()


def tool_get_slide_count(args_str: str) -> str:
    return get_slide_count()


def tool_search_topic(query: str) -> str:
    return search_topic(query.strip())


# ──────────────────────────────────────────────────────────────────────────────
# SYSTEM PROMPT
# ──────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are Auto-PPT, an expert presentation agent.
Your job: turn a user's single-sentence request into a polished PowerPoint file.

User Input -> Agent Brain (LLM) -> Decision: "Write Slide 1"
             -> Tool Call (MCP: write_slide)
             -> Agent Brain -> Decision: "Write Slide 2"
             -> Tool Call...
             -> Agent Brain -> Decision: "Save File" -> FINISH

MANDATORY 4-STEP AGENTIC LOOP:

STEP 1 - PLAN (MUST be done before ANY tool call)
  Think through: audience, tone, number of slides requested.
  Produce a JSON outline in your Thought like:
  {{
    "output_file": "workspace/stars_6th_grade.pptx",
    "slides": [
      {{"index": 0, "type": "title",   "title": "...", "subtitle": "..."}},
      {{"index": 1, "type": "content", "title": "...", "bullets": ["...", "...", "..."]}},
      {{"index": N, "type": "summary", "title": "Key Takeaways", "bullets": ["...", "..."]}}
    ]
  }}

STEP 2 - RESEARCH
  For each content slide topic, call search_topic ONCE to gather real facts.
  If search_topic fails, use your own knowledge - NEVER crash.

STEP 3 - BUILD
  a. Call create_presentation with the output file path.
  b. For each slide call add_slide with correct type, title, bullets.
     - title slides   : {{"slide_type":"title",   "title":"...", "subtitle":"..."}}
     - content slides : {{"slide_type":"content", "title":"...", "bullets":["...","...","..."]}}
     - summary slides : {{"slide_type":"summary", "title":"...", "bullets":["...","..."]}}

STEP 4 - SAVE
  Call save_presentation {{}}.
  Then reply: "Done! Presentation saved."

RULES:
- NEVER skip Step 1 planning.
- ALWAYS have at least 1 title slide and 1 summary slide.
- Bullet points must be factual and audience-appropriate.
- If the user says "5 slides", produce exactly 5.
- Pass tool arguments as valid JSON strings.
- If a tool returns an error, fix arguments and retry once.



You have access to these tools:
{tools}

Tool names you can use: {tool_names}

Use this format EXACTLY:

Thought: <your reasoning>
Action: <tool name>
Action Input: <JSON string>
Observation: <tool result>
... (repeat as needed)
Thought: I now know the final answer.
Final Answer: <summary message>

Begin!

Question: {input}
{agent_scratchpad}"""


# ──────────────────────────────────────────────────────────────────────────────
# MAIN AGENT RUNNER
# ──────────────────────────────────────────────────────────────────────────────

def run_ppt_agent(user_request: str) -> str:
    """Build and run the LangChain ReAct agent."""

    tools = [
        Tool(name="create_presentation",
             func=tool_create_presentation,
             description=(
                 "Initialise a new .pptx file. MUST be called first. "
                 'Input JSON: {"output_path": "workspace/filename.pptx"}'
             )),
        Tool(name="add_slide",
             func=tool_add_slide,
             description=(
                 "Add one slide. "
                 'Input JSON: {"slide_type": "title|content|summary", '
                 '"title": "...", "subtitle": "...", "bullets": ["...","..."]}'
             )),
        Tool(name="write_text",
             func=tool_write_text,
             description=(
                 "Overwrite text on existing slide (0-indexed). "
                 'Input JSON: {"slide_index": 0, "title": "...", "bullets": ["..."]}'
             )),
        Tool(name="add_image_placeholder",
             func=tool_add_image_placeholder,
             description=(
                 "Add a visual placeholder box on a slide. "
                 'Input JSON: {"slide_index": 1, "label": "[Image: ...]", '
                 '"position": "left|right|center"}'
             )),
        Tool(name="save_presentation",
             func=tool_save_presentation,
             description="Save the presentation to disk. Call this LAST. Input JSON: {}"),
        Tool(name="get_slide_count",
             func=tool_get_slide_count,
             description="Returns current number of slides. Input JSON: {}"),
        Tool(name="search_topic",
             func=tool_search_topic,
             description=(
                 "Search the web for facts about a topic. "
                 "Input: plain text query string (NOT JSON). "
                 "Example: life cycle of a star facts for kids"
             )),
    ]

    # ── Load Gemini LLM ──────────────────────────────────────────────────────
    logger.info(f"Loading model: {MODEL_ID} ...")
    llm = ChatGoogleGenerativeAI(
        model=MODEL_ID,
        temperature=TEMPERATURE,
        max_output_tokens=MAX_TOKENS,
        google_api_key=os.getenv("GOOGLE_API_KEY"),
    )
    logger.info("✅ Model loaded!")

    # ── Build ReAct Agent ─────────────────────────────────────────────────────
    prompt = PromptTemplate.from_template(SYSTEM_PROMPT)
    agent  = create_react_agent(llm=llm, tools=tools, prompt=prompt)

    agent_executor = AgentExecutor(
        agent=agent,
        tools=tools,
        verbose=AGENT_VERBOSE,
        max_iterations=AGENT_MAX_ITER,
        handle_parsing_errors=True,
        return_intermediate_steps=False,
    )

    logger.info("🤖 Agent starting...")
    result = agent_executor.invoke({"input": user_request})
    return result.get("output", "Agent finished with no output.")


# ──────────────────────────────────────────────────────────────────────────────
# CLI ENTRY POINT
# ──────────────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) > 1:
        user_request = " ".join(sys.argv[1:])
    else:
        print("Auto-PPT Agent")
        print("=" * 50)
        user_request = input("Enter your presentation request:\n> ").strip()
        if not user_request:
            user_request = "Create a 5-slide presentation on the life cycle of a star for a 6th-grade class"
            print(f"Using default: {user_request}")

    print(f"\n🚀 Running agent for: '{user_request}'\n")
    result = run_ppt_agent(user_request)
    print(f"\n{'='*50}")
    print(f"✅ Agent Output:\n{result}")


if __name__ == "__main__":
    main()