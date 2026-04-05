# Reflection Document – Auto-PPT Agent

**Course:** AI Agents & MCP Architecture  
**Assignment:** Auto-PPT Agent

---

## Where Did the Agent Fail Its First Attempt?

The very first thing that broke was something I didn't even think about: **argument parsing**. When LangChain's ReAct agent calls a tool, it passes the arguments as a plain text string. But my MCP tools were expecting a proper Python dictionary. So when the agent tried to call `add_slide`, it passed something like `title='Life Cycle of a Star', slide_type='title'` with single quotes — which is not valid JSON. The whole thing crashed immediately with a `JSONDecodeError`. I had to build a small parser that tries `json.loads()` first, then fixes single quotes, and finally falls back to `eval()` as a last resort.

The second failure is the agent just... didn't do anything. It would read the prompt, think for a second, and then reply "Here is a 5-slide presentation on stars" — as a chat message. No tools called, no file created, nothing. It was basically just a chatbot pretending to be an agent. The fix was rewriting the entire system prompt to force a **mandatory 4-step loop**: Plan first, then Research, then Build, then Save. Once the prompt made it crystal clear that the agent must not skip Step 1 or answer directly, it finally started behaving like an actual agent.

The third failure was the agent inventing tool names. It would confidently call `create_slide` when the actual tool was called `add_slide`, or `finalize` instead of `save_presentation`. LangChain would throw a `ValueError: Tool not found` and everything stopped. Enabling `handle_parsing_errors=True` helped a lot here — instead of crashing, the agent would receive the error as feedback and try to correct itself.

---

## How Did MCP Prevent Me From Writing Hardcoded Scripts?

Before I understood MCP, my instinct was to just write a single Python script that creates the slides directly — something like:

```python
prs = Presentation()
slide1 = prs.slides.add_slide(layout)
slide1.shapes.title.text = "Life Cycle of a Star"
prs.save("output.pptx")
```

That works perfectly fine — for exactly one topic, one slide count, one structure. The moment someone asks for a different topic, the whole script needs to be rewritten.

MCP changed this completely. By putting all the presentation logic inside a server with named tools like `add_slide` and `save_presentation`, the agent only knows *what* to call — not *how* it works underneath. The actual `python-pptx` drawing code, the colour themes, the font sizes — all of that is hidden inside the server. The agent just says "add a content slide about nebulae with these bullets" and the server handles the rest.

What I found most powerful was that this separation meant I could improve the slide design (add the Midnight Star theme, better layouts, gold bullet markers) without touching the agent code at all. The agent kept working exactly the same. That's the real value of MCP — it creates a clean boundary between the thinking and the doing.

---

## Every Error and Challenge I Faced 

This project had way more problems than I expected. Here's what actually happened, in the order I hit them:

---

**The Python version problem** came first. I already had Python 3.14 installed, which turned out to be too new for almost every package I needed. None of the wheels existed for it yet. I had to install Python 3.11.5 separately and create a virtual environment that specifically used 3.11. I got to know that Lang chain is much compatible with python 3.11 version

**A fake package in requirements.txt** was embarrassing. I had `asyncio-compat>=0.1.0` listed as a dependency. It doesn't exist. `asyncio` is just built into Python. Had to remove it and install everything else manually.

**LangChain version conflicts** turned into a multi-hour rabbit hole. The first install gave me `langchain 1.2.15`, which had removed `AgentExecutor` from the location my code expected. I tried to downgrade, but that broke `langchain-huggingface` and `langgraph-prebuilt` in a circular dependency loop that just kept getting worse. In the end I deleted the entire virtual environment and started fresh with exact pinned version numbers — that was the only clean way out.

**The MCP crash on Windows** was genuinely puzzling. I had written the agent to spin up the MCP servers as subprocesses using stdio transport — which is how MCP is supposed to work. But on Windows, the asyncio event loop handles subprocess communication differently, and the MCP session initialization would hang and then crash with a `CancelledError` and `RuntimeError: Attempted to exit cancel scope in a different task`. After a lot of digging, the fix was surprisingly simple: since my MCP tool functions decorated with `@mcp.tool()` are also plain Python functions, I could just import them directly instead of going through the subprocess protocol. The MCP architecture was still there — the server files, the decorators, the tool schemas — but the actual calling happened directly.

**A type hint that worked everywhere except my machine.** The line `_prs: Presentation | None = None` crashed with `TypeError: unsupported operand type(s) for |: 'function' and 'NoneType'`. It turns out `python-pptx`'s `Presentation` is internally a callable/function wrapper, so Python couldn't use the `|` union syntax on it. Switching to `Optional[Presentation]` from the `typing` module fixed it.

**The save function crashing on a blank string** was a subtle one. When the output path is just `"output.pptx"` with no folder, `os.path.dirname("output.pptx")` returns an empty string `""`. Then `os.makedirs("")` crashes with "No such file or directory". The fix was converting the path to an absolute path with `os.path.abspath()` first, which gives you a full path you can safely extract the directory from.

**HuggingFace parameter errors** came in two waves. First, I had `max_retries=3` in the `HuggingFaceEndpoint` config — which is a valid parameter for OpenAI clients but not for HuggingFace. LangChain didn't reject it, it just silently passed it deeper into the inference client, which then crashed. After removing that, I hit a `StopIteration` error because `huggingface_hub 1.9.0` introduced "inference provider mapping" — and the models I was trying to use (Phi-3.5-mini and Qwen2.5) simply didn't have free inference providers registered anymore. That's when I switched to Google Gemini, which has no such mapping requirement and a much more generous free tier.

**Gemini had its own quota issue.** The first model I tried gave a `Quota exceeded... limit: 0` error — the model existed but wasn't enabled for my specific Google AI Studio project. Switched to a model that showed healthy RPM quotas in the dashboard and it worked immediately.

**A missing variable crashed the whole agent setup** before it even ran a single step. LangChain's `create_react_agent` requires both `{tools}` and `{tool_names}` in the prompt template. I had `{tools}` but forgot `{tool_names}`. It threw `ValueError: Prompt missing required variables: {'tool_names'}` and refused to start.

**DuckDuckGo rate limiting** hit during testing. The search tool was returning `202 Ratelimit` errors because I was making too many requests too quickly. Added a random 1.5–3 second delay between requests and built two fallback layers — DuckDuckGo news and Wikipedia — so the agent always gets something useful even if the main search fails.

**The agent being "lazy"** was a separate problem I hit even after fixing the tool-calling format. The LLM would read the presentation request and just generate a nicely formatted text response describing what the slides would look like — completely ignoring all the tools. The system prompt needed to be much more explicit: no answering directly, no skipping the planning step, every response must follow the Thought/Action/Observation format until the file is saved.

**JSON formatting failures mid-run** happened occasionally where the LLM would produce something like `Missing 'Action:' after 'Thought:'` — it would think correctly but then output the action in the wrong format. `handle_parsing_errors=True` in the AgentExecutor catches these and feeds the error back to the model as an observation, giving it a chance to self-correct.

**The PowerPoint looking plain** was something I only noticed after downloading and opening the file. The Streamlit preview looked great because of the CSS styling — but the actual `.pptx` was just white slides with black text and default bullet points. I had to completely rewrite the `pptx_mcp_server.py` to manually draw the Midnight Star theme using `python-pptx`'s shape and colour APIs — navy backgrounds, gold accent bars, custom bullet dots, three distinct slide layouts.

**Duplicate content in the preview** appeared after adding the design theme. The slide preview was showing "SLIDE 4" and "CONTENT" labels twice because the `extract_slides()` function was reading the decorative badge shapes from the actual `.pptx` file and treating them as real slide content. Fixed by filtering out short strings that match known decorative patterns.

**Windows encoding crashing the servers** on startup was unexpected — the emoji characters in print statements couldn't be displayed by Windows' default `cp1252` terminal encoding. Added `sys.stdout.reconfigure(encoding='utf-8')` to the server files.

**The Streamlit example chips not working** was frustrating because they looked fine but did nothing. Clicking "Create a 5-slide presentation on stars…" should fill the text box — but `st.session_state["prompt_input"] = ex` doesn't work once the widget has already been rendered. The fix was using a separate `prompt_value` state variable that drives the `value=` parameter of the text area instead of trying to modify the widget directly.

**The Streamlit startup crash** was the very first error users would see when opening the app. `st.text_area()` returns `None` before the user has typed anything, so calling `.strip()` on it immediately crashed with `AttributeError: 'NoneType' object has no attribute 'strip'`. Fixed by guarding every usage with `(user_prompt or "").strip()`.

---

## The Biggest Lesson

If I had to pick one thing this project taught me, it's that the hardest part of building an AI agent isn't the AI — it's all the plumbing around it. Getting the LLM to reason correctly is maybe 20% of the work. The other 80% is error handling, format parsing, dependency management, platform-specific bugs, and making sure every layer of the system speaks the same language as the next.

MCP was genuinely useful here because it drew a hard line between the agent's reasoning and the tools it uses. When something broke, I always knew which side of that line the problem was on. That made debugging much faster than it would have been with everything tangled together in one script.
