# 🪄 Auto-PPT Agent – AI-Driven Presentation Engine

> **Generate research-backed, professionally styled, and perfectly aligned PowerPoint presentations from a single prompt.**

[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![Gemini 3.1](https://img.shields.io/badge/Model-Gemini%203.1%20Flash%20Lite-orange.svg)](https://ai.google.dev/)
[![LangChain](https://img.shields.io/badge/Agent-LangChain%20ReAct-green.svg)](https://www.langchain.com/)
[![MCP Architecture](https://img.shields.io/badge/Architecture-MCP-blue.svg)](https://modelcontextprotocol.io/)

---

## 📽️ Showcase & Demo
[![Watch the Demo](https://img.youtube.com/vi/YOUR_VIDEO_ID_HERE/maxresdefault.jpg)](https://drive.google.com/drive/folders/1IJmTzRCD9qXs3-V_WXXVNWaQ_YBw3zwP?usp=drive_link)
> **Replace `YOUR_VIDEO_ID_HERE` with your actual video link to showcase the agent in action!**

---

## 🧐 What is Auto-PPT Agent?
Auto-PPT Agent is an autonomous AI assistant that bridging the gap between raw prompts and presentation-ready decks. Unlike static templates, it **researches**, **plans**, and **builds** each slide from scratch using real-time web data and custom styling logic.

---

## 🧱 Repository Structure
```text
auto_ppt_agent/
├── app.py                # Streamlit Web Frontend – The visual interface
├── agent_ppt.py          # Core Agent Brain – LangChain ReAct orchestration
├── pptx_mcp_server.py    # MCP Server (Hands) – Premium PPTX drawing tools
├── search_mcp_server.py  # MCP Server (Eyes) – Real-time web research
├── .env.example          # Environment Config – Secure credentials & model settings templete
├── .gitignore            # Git Exclusions – Prevents sensitive file leakage
├── requirements.txt      # Dependencies – Project-specific libraries
├── workspace/            # Project Output – Where your .pptx files are saved
└── README.md             # Documentation – You are here!
```

---

## 🏗️ Architecture Flow
The project follows a **Brain (LLM)** & **Hands (MCP)** architecture:

### 1. The Planning Phase (Thought)
The **Gemini 3.1** engine receives your prompt and generates a multi-slide JSON plan, ensuring a logical flow from Title to Summary.

### 2. The Research Phase (Observation)
For each content slide, the agent calls the **Search MCP Server** to fetch up-to-date facts from the web via DuckDuckGo, ensuring high accuracy.

### 3. The Building Phase (Action)
The **PPTX MCP Server** draws the presentation using a "Midnight Star" design system. Each slide is built with a 60/40 split-layout for text and images.

### 4. The Final Review
The agent verifies all slides are present, saves the file to the `./workspace/` directory, and provides a direct download link of the output.

---

## 🛠️ Features
*   **📐 Premium Split-Layout**: Every content slide includes an automatic image placeholder aligned to the right.
*   **🎨 Midnight Star Theme**: Elegant deep-navy background with gold accents and tri-color badge indicators.
*   **🌐 Real-Time Research**: No more outdated information; the agent researches every topic live.
*   **🖥️ Dashboard Interface**: A modern Streamlit UI with real-time "Thinking" logs and live slide previews.
*   **🤖 Multi-Tool Support**: Integrated Planning, Searching, Building, and Saving tools.

---

## 🚀 Setup & Installation

### 1. Prerequisites
*   Python 3.11 or higher
*   Google Gemini API Key ([Get one here](https://aistudio.google.com/))

### 2. Installation
```bash
# Clone the repository
git clone https://github.com/yourusername/auto-ppt-agent.git
cd auto-ppt-agent

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

### 3. Configuration
Rename the `.env` placeholder or create a new one:
```env
GOOGLE_API_KEY=AIzaSy...
MODEL_ID=gemini-3.1-flash-lite-preview
MAX_TOKENS=512
TEMPERATURE=0.2
AGENT_MAX_ITERATIONS=20
```

---

## 📈 Configuration Details

| Variable | Default | Purpose |
| :--- | :--- | :--- |
| `MODEL_ID` | `gemini-3.1-flash-lite-preview` | The LLM "Brain" to use. |
| `AGENT_MAX_ITERATIONS` | `20` | Max steps the agent can take per run. |
| `AGENT_WORKSPACE` | `./workspace` | Directory for generated presentations. |
| `AGENT_VERBOSE` | `True` | Show detailed "Thought" logs in terminal. |

---

## 🕹️ Usage

### Option A: Web Dashboard (Recommended)
```bash
streamlit run app.py
```
Enjoy a visual builder with real-time slide previews!

### Option B: Command Line
```bash
python agent_ppt.py "Create a 5-slide deck on Quantum Computing for beginners"
```

---
