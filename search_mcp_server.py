"""
search_mcp_server.py
--------------------
MCP Server #2 — Web Search Tool

Uses FastMCP with @mcp.tool() decorator.

Search strategy (tries in order):
  1. DuckDuckGo text search  (ddgs.text)
  2. DuckDuckGo news search  (ddgs.news)  ← fallback if text is rate limited
  3. Wikipedia summary        ← fallback if both DDG methods fail
  4. Graceful message         ← agent uses its own knowledge

Tools:
  - search_topic : Search the web and return text snippets

Run:
  python search_mcp_server.py
"""

import time
import random

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("search-mcp-server")


# ──────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────────────────────────────────────

def _ddg_text_search(query: str, max_results: int) -> list[str]:
    """Try DuckDuckGo text search with a small delay to avoid rate limits."""
    from duckduckgo_search import DDGS
    time.sleep(random.uniform(1.5, 3.0))   # polite delay
    results = []
    with DDGS() as ddgs:
        for r in ddgs.text(query, max_results=max_results):
            snippet = r.get("body", "").strip()
            title   = r.get("title", "").strip()
            if snippet:
                results.append(f"[{title}]\n{snippet}")
    return results


def _ddg_news_search(query: str, max_results: int) -> list[str]:
    """Fallback: DuckDuckGo news search."""
    from duckduckgo_search import DDGS
    time.sleep(random.uniform(1.5, 3.0))
    results = []
    with DDGS() as ddgs:
        for r in ddgs.news(query, max_results=max_results):
            snippet = r.get("body", "").strip()
            title   = r.get("title", "").strip()
            if snippet:
                results.append(f"[{title}]\n{snippet}")
    return results


def _wikipedia_search(query: str) -> list[str]:
    """Fallback: fetch Wikipedia summary for the topic."""
    import urllib.request
    import urllib.parse
    import json

    search_term = urllib.parse.quote(query)
    url = (
        f"https://en.wikipedia.org/api/rest_v1/page/summary/"
        f"{search_term.replace('%20', '_')}"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "AutoPPTAgent/1.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read().decode())
        extract = data.get("extract", "")
        title   = data.get("title", query)
        if extract:
            return [f"[Wikipedia: {title}]\n{extract}"]
    return []


# ──────────────────────────────────────────────────────────────────────────────
# MCP TOOL
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def search_topic(query: str, max_results: int = 5) -> str:
    """
    Search the web for factual information on a topic.
    Tries DuckDuckGo text → DuckDuckGo news → Wikipedia in order.
    Falls back gracefully if all methods fail — never crashes the agent.

    Args:
        query       : Search query e.g. 'life cycle of a star facts for kids'
        max_results : Number of result snippets to return (1-10). Default 5.
    """
    max_results = max(1, min(int(max_results), 10))
    results     = []
    errors      = []

    # ── Method 1: DuckDuckGo text ─────────────────────────────────────────────
    try:
        results = _ddg_text_search(query, max_results)
        if results:
            combined = "\n\n".join(results)
            return f"Search results for '{query}':\n\n{combined}"
    except Exception as e:
        errors.append(f"DDG text: {e}")

    # ── Method 2: DuckDuckGo news ─────────────────────────────────────────────
    try:
        results = _ddg_news_search(query, max_results)
        if results:
            combined = "\n\n".join(results)
            return f"Search results (news) for '{query}':\n\n{combined}"
    except Exception as e:
        errors.append(f"DDG news: {e}")

    # ── Method 3: Wikipedia ───────────────────────────────────────────────────
    try:
        results = _wikipedia_search(query)
        if results:
            combined = "\n\n".join(results)
            return f"Wikipedia summary for '{query}':\n\n{combined}"
    except Exception as e:
        errors.append(f"Wikipedia: {e}")

    # ── Graceful fallback ─────────────────────────────────────────────────────
    return (
        f"⚠️ All search methods failed for '{query}'. "
        f"Errors: {'; '.join(errors)}. "
        f"Please use your own knowledge to generate accurate content for this topic."
    )


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    mcp.run()