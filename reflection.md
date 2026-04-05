# 📔 Auto-PPT Agent – Reflection Document

> **Assignment:** AI Agents & MCP Architecture  
> **Course:** AI Agents & MCP Architecture  
>  Varshini  

---

## 1. Where did your agent fail its first attempt?
My agent faced two primary challenges during its initial development phase:

### Challenge A: Sequential vs. Strategic Thinking
**The Failure:** In my first version, the agent attempted to write slides one after another without a global plan. This resulted in a fragmented presentation where the story didn't flow correctly (e.g., Slide 3 would introduce a topic that Slide 1 had already covered).
**The Fix:** I redesigned the **System Prompt** to enforce a mandatory 4-step agentic loop: **Plan → Research → Build → Save**. By forcing the agent to generate a full JSON outline *before* calling any slide tools, I ensured a logical and coherent narrative flow.

### Challenge B: Model Access & Rate Limiting
**The Failure:** The initial implementation used `gemini-2.0-flash`, which hit a "Limit 0" quota error in my specific AI Studio project.
**The Fix:** I migrated the configuration to use **Gemini 3.1 Flash Lite**. This model provided a more generous quota (15 RPM) while maintaining excellent tool-calling accuracy for the complex multi-step PowerPoint generation process.

---

## 2. How did MCP prevent you from writing hardcoded scripts?
The **Model Context Protocol (MCP)** was a game-changer because it decoupled the **Agent's Brain** (the LLM) from the **Agent's Hands** (the PPTX tools).

### De-coupling via abstraction
Instead of writing a traditional, hardcoded script where every font size and color was buried in a giant main loop, I created a dedicated **FastMCP Server** (`pptx_mcp_server.py`). 
*   **The "Hands"**: The MCP server handles the complex layout logic, theme management ("Midnight Star"), and split-layout image placeholders.
*   **The "Brain"**: The agent only needs to know *what* to put on a slide, not *how* to draw it.

Because of MCP, I replaced hundreds of lines of hardcoded `python-pptx` logic with a simple set of tools: `create_presentation`, `add_slide`, and `write_text`. This makes the system **extensible**—if I want to change the entire visual theme tomorrow, I only update the MCP server, and the agent automatically adapts without any changes to its core reasoning logic.

---

## 3. Reflection on Robustness & Planning
By using a **3-tier research strategy** in the `search_mcp_server`, the agent handles vague or niche prompts gracefully. If DuckDuckGo yields no results, it falls back to Wikipedia, and if that fails, it relies on its internal knowledge (graceful hallucination) as per the assignment requirements. This multi-layered architecture is only possible because of the **structured MCP tool definition**.

---

**Conclusion:** The combination of a robust LangChain ReAct executor and specialized MCP servers transformed a simple script into a sophisticated, autonomous presentation engine.
