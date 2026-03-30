
# Agent Memory Engineering Knowledge Base
Source basis: attached transcript (`Pasted text.txt`) and notebooks (`L2.ipynb`, `L3.ipynb`, `L4.ipynb`, `L5.ipynb`).

This document is written to be ingested by an LLM as a project knowledge base. It is intentionally structured, implementation-heavy, and oriented toward reusable patterns rather than high-level course summary.

---

## 1. Core Thesis

A production agent should not rely only on prompt history. It should use **externalized memory** backed by a database, with multiple memory stores optimized for different retrieval and update patterns.

The architecture in the transcript and notebooks uses:

- an **LLM** for reasoning and planning
- an **embedding model** for semantic encoding
- a **database** as the external memory core
- a **MemoryManager** as the abstraction layer over storage and retrieval
- a **Toolbox** as retrievable procedural memory for tool selection
- **summary memory** for context reduction and recovery
- **workflow memory** for storing successful step sequences
- **entity memory** for named references and continuity

The transcript explicitly frames the database as the **Agent Memory Core** and treats memory engineering as the discipline required to turn stateless LLMs into persistent, stateful, adaptive agents.

---

## 2. Agent Capability Model

The notebooks and transcript define an agent as a system with four capabilities:

1. **Perception** — accept user input and environment observations
2. **Reasoning** — use an LLM to interpret state and decide next steps
3. **Action** — call tools, APIs, or local functions
4. **Memory** — store and retrieve information across turns and sessions

Without memory, the agent is only a single-turn tool router.
With memory, it becomes persistent.
With memory awareness, it can reason about its own memory stores, summaries, workflows, and tool affordances.

---

## 3. Progression of Agent Designs

### 3.1 Stateless agent
A stateless agent only sees the current turn.
It can answer direct questions but fails at follow-ups, cross-session continuity, or long-running tasks.

Example:
- User: "Recommend restaurants nearby."
- Agent: returns a ranked list.
- User: "Book the first one."
- Stateless agent fails because "the first one" depends on earlier state.

### 3.2 Memory-augmented agent
A memory-augmented agent stores prior interactions in an external memory store and retrieves relevant context later.

Example:
- It stores prior restaurant recommendations.
- When asked to book the first one, it retrieves the prior list and resolves the reference.

### 3.3 Memory-aware agent
A memory-aware agent is not only connected to memory, but **aware of memory types and their roles**.
The agent system prompt teaches the model how to use:

- Conversation Memory
- Knowledge Base Memory
- Workflow Memory
- Entity Memory
- Summary Memory

It can also call memory operations as tools, such as:
- `expand_summary(summary_id)`
- `summarize_and_store(text, thread_id=...)`

This is the final pattern implemented in `L5.ipynb`.

---

## 4. Memory Taxonomy

## 4.1 Working memory
The active LLM context window and temporary scratchpad used during reasoning.
It is ephemeral and usually lost after the request or session.

Used for:
- current prompt
- intermediate tool results
- short-lived plan state

## 4.2 Semantic cache
A cache of previously answered semantically similar queries.
Uses embeddings and vector lookup to avoid recomputation.

Used for:
- repeated user questions
- low-latency assistants
- FAQ-style retrieval

## 4.3 Conversational memory (episodic)
Stores user/assistant interaction history as ordered records with timestamps.

Used for:
- follow-up understanding
- thread continuity
- user preferences
- unresolved tasks

Typical fields:
- id
- thread_id
- role
- content
- timestamp
- metadata
- created_at
- summary_id

## 4.4 Knowledge base memory (semantic)
Stores chunked external information with embeddings and metadata.
This is the durable fact memory for factual grounding.

Used for:
- RAG
- technical manuals
- papers
- product docs
- web search results stored for later use

## 4.5 Workflow memory (procedural)
Stores prior successful sequences of actions, outcomes, or execution traces.

Used for:
- repeated workflows
- long-horizon tasks
- planning based on previous successful runs
- recovering from partial failures

## 4.6 Toolbox memory (procedural/action memory)
Stores tool definitions as retrievable memory units rather than stuffing all tools into every prompt.

Used for:
- large tool ecosystems
- API-heavy agents
- semantic tool routing
- scaling tool use without context bloat

## 4.7 Entity memory
Stores named entities, references, and descriptors extracted from user queries or outputs.

Used for:
- person, place, org, project, paper, model, API references
- disambiguation
- referential continuity across sessions

## 4.8 Summary memory
Stores compressed summaries of old conversation or context plus an ID that can later be expanded.

Used for:
- context compaction
- long-running threads
- bounded context windows
- recovering details later with a tool call

---

## 5. Why Multiple Memory Stores Matter

A single giant transcript is a poor memory design.

Different memory types require different storage and retrieval behavior:

- Conversation memory needs **time order** and **thread scoping**
- Knowledge memory needs **semantic similarity**
- Workflow memory needs **pattern recall**
- Toolbox memory needs **retrieval over tool affordances**
- Entity memory needs **reference tracking**
- Summary memory needs **compression plus recoverability**

This is why the implementation uses:
- SQL tables for conversational and tool logs
- vector stores for knowledge, workflow, toolbox, entity, and summary stores

---

## 6. Memory Units

A **memory unit** is the smallest atomic stored unit used by the agent.

Examples:

### Conversational memory unit
A single row:
- thread_id
- role
- content
- timestamp
- optional metadata
- optional summary_id

### Workflow memory unit
A single row:
- query or task
- steps taken
- final outcome
- timestamp
- vector embedding for similarity retrieval

### Tool memory unit
A single row:
- tool name
- tool description
- parameter schema
- optionally LLM-augmented description
- embedding

### Summary memory unit
A single row:
- summary_id
- raw/original text
- summary text
- description label
- embedding

Memory units matter because retrieval quality depends on **how you serialize and store one unit of memory**.

---

## 7. Data Layer / Memory Layer Architecture

The notebooks compress the broader agent stack into three layers:

1. **Application Layer**
   - user interaction
   - orchestration
   - agent loop

2. **Memory Layer**
   - memory core (database)
   - memory manager
   - vector stores
   - conversation store
   - summary store
   - toolbox store

3. **Infrastructure Layer**
   - Oracle database
   - embedding model
   - LLM provider
   - indexing mechanisms

The **Memory Layer** is the most important concept for project implementation.
It acts like a dedicated memory subsystem.

---

## 8. L2 Notebook: Constructing the Memory Manager

`L2.ipynb` is the infrastructure notebook.
It sets up database connections, embedding model, tables, vector stores, indexes, and the `MemoryManager`.

### 8.1 Conversational history table

The notebook creates a relational SQL table for conversation history.

```python
def create_conversational_history_table(conn, table_name: str = "CONVERSATIONAL_MEMORY"):
    with conn.cursor() as cur:
        try:
            cur.execute(f"DROP TABLE {table_name}")
        except:
            pass

        cur.execute(f"""
            CREATE TABLE {table_name} (
                id VARCHAR2(100) DEFAULT SYS_GUID() PRIMARY KEY,
                thread_id VARCHAR2(100) NOT NULL,
                role VARCHAR2(50) NOT NULL,
                content CLOB NOT NULL,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                metadata CLOB,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                summary_id VARCHAR2(100) DEFAULT NULL
            )
        """)

        cur.execute(f"""
            CREATE INDEX idx_{table_name.lower()}_thread_id ON {table_name}(thread_id)
        """)

        cur.execute(f"""
            CREATE INDEX idx_{table_name.lower()}_timestamp ON {table_name}(timestamp)
        """)

    conn.commit()
    return table_name
```

### Why this is used
This schema is designed for:
- thread-level retrieval
- role-aware message reconstruction
- chronological ordering
- summary compaction tracking via `summary_id`

### In your agent project
Use this exact idea for:
- chat history persistence
- thread reconstruction
- follow-up resolution
- linking summarized messages to a later summary unit

---

## 8.2 StoreManager abstraction

The notebook defines `StoreManager` to hold all vector stores.

```python
class StoreManager:
    def __init__(self, client, embedding_function, table_names, distance_strategy,
                 conversational_table, tool_log_table: str | None = None):
        self.client = client
        self.embedding_function = embedding_function
        self.distance_strategy = distance_strategy
        self._conversational_table = conversational_table
        self._tool_log_table = tool_log_table

        self._knowledge_base_vs = OracleVS(
            client=client,
            embedding_function=embedding_function,
            table_name=table_names['knowledge_base'],
            distance_strategy=distance_strategy,
        )

        self._workflow_vs = OracleVS(
            client=client,
            embedding_function=embedding_function,
            table_name=table_names['workflow'],
            distance_strategy=distance_strategy,
        )

        self._toolbox_vs = OracleVS(
            client=client,
            embedding_function=embedding_function,
            table_name=table_names['toolbox'],
            distance_strategy=distance_strategy,
        )

        self._entity_vs = OracleVS(
            client=client,
            embedding_function=embedding_function,
            table_name=table_names['entity'],
            distance_strategy=distance_strategy,
        )

        self._summary_vs = OracleVS(
            client=client,
            embedding_function=embedding_function,
            table_name=table_names['summary'],
            distance_strategy=distance_strategy,
        )
```

### Why this is used
This is a clean separation of stores by memory semantics.
Instead of one vector table for everything, the design uses multiple vector stores:
- semantic memory
- workflow memory
- toolbox memory
- entity memory
- summary memory

### In your agent project
If you already have vector DB infrastructure, reproduce the pattern even if you do not use Oracle.
The important part is **logical separation by memory type**, not Oracle specifically.

---

## 8.3 MemoryManager initialization

The notebook then wraps all stores in a `MemoryManager`.

```python
memory_manager = MemoryManager(
    conn=database_connection,
    conversation_table=CONVERSATION_HISTORY_TABLE, 
    knowledge_base_vs=knowledge_base_vs,
    workflow_vs=workflow_vs,
    toolbox_vs=toolbox_vs,
    entity_vs=entity_vs,
    summary_vs=summary_vs,
    tool_log_table=TOOL_LOG_HISTORY_TABLE
)
```

### Why this is used
`MemoryManager` is the unified API for memory operations:
- `write_conversational_memory`
- `read_conversational_memory`
- `write_knowledge_base`
- `read_knowledge_base`
- `write_workflow`
- `read_workflow`
- `write_entity`
- `read_entity`
- `write_summary`
- `read_summary_memory`
- `read_summary_context`
- `write_tool_log`
- `read_toolbox`

### Design lesson
Your agent should not talk directly to raw storage from the main agent loop.
Instead:
- one layer owns memory semantics
- agent loop calls memory operations through that layer

This is essential if you want to evolve memory behavior later without rewriting the agent loop.

---

## 8.4 Knowledge base ingestion example

`L2.ipynb` ingests papers into semantic memory.

```python
for paper in islice(ds, 100):
    title = (paper.get("title") or "").strip()
    abstract = (paper.get("abstract") or "").strip()
    subjects = (paper.get("subjects") or paper.get("primary_subject") or "").strip()
    submission_date = (paper.get("submission_date") or "").strip()

    if not (title or abstract or subjects):
        continue

    text = "\n".join([part for part in (title, subjects, abstract) if part])

    memory_manager.write_knowledge_base(
        text=text,
        metadata={
            "arxiv_id": paper.get("arxiv_id"),
            "title": title,
            "subjects": subjects,
            "abstract": abstract,
            "submission_date": submission_date,
        },
    )
```

### Why this is used
This converts raw data into semantic memory units.
Each row contains:
- text for embedding
- metadata for attribution and filtering

### In your project
This is the pattern for adding:
- docs
- tickets
- logs
- papers
- web search results
- code explanations
- internal design docs

A good memory write format for KB units is:

```python
memory_manager.write_knowledge_base(
    text=chunk_text,
    metadata={
        "source_type": "research_paper",
        "source_id": "...",
        "title": "...",
        "tags": ["memory", "rag", "agents"],
        "ingested_at": "...",
        "chunk_id": 3,
        "num_chunks": 18
    }
)
```

---

## 9. L3 Notebook: Semantic Tool Memory / Toolbox Pattern

`L3.ipynb` solves the tool scaling problem.

### Problem
If you stuff every tool definition into every prompt:
- token cost rises
- latency rises
- context gets noisy
- tool selection degrades

### Solution
Store tool definitions as memory units and retrieve only relevant tools for the query.

---

## 9.1 Toolbox lookup tool

```python
@toolbox.register_tool(augment=True)
def read_toolbox(query: str, k: int = 3) -> list[str]:
    """
    Search the toolbox for functions that can help solve a problem or complete a task.
    """
    return memory_manager.read_toolbox(query, k=k)
```

### Why this is used
This gives the agent a meta-tool to discover capabilities based on semantic intent.

### Example use
If the agent is facing:
- a web search need
- a need to fetch papers
- a need to expand summary memory
- a need to summarize old context

It can query toolbox memory instead of being forced to know every tool ahead of time.

### In your project
If your agent has 50+ functions, use a toolbox retrieval stage before model inference:
1. embed user query
2. retrieve top-k tools
3. pass only those tool schemas to the model

---

## 9.2 Why `augment=True` matters
Tool registration often uses `augment=True`.
That means the system enriches the tool definition using an LLM before embedding it.

This improves:
- retrieval recall
- semantic separability
- richer affordance description

### Example
Raw docstring:
- "Returns current time"

Augmented tool description:
- explains when to call it
- lists parameters
- lists return type
- clarifies use cases
- adds step-by-step semantics

This produces better embeddings than a tiny raw docstring.

---

## 9.3 Tavily search tool with write-back to KB

```python
from tavily import TavilyClient
from datetime import datetime

tavily_client = TavilyClient()

@toolbox.register_tool(augment=True)
def search_tavily(query: str, max_results: int = 5):
    """
    Use this function to search the web and store the results in the knowledge base.
    """
    response = tavily_client.search(query=query, max_results=max_results)
    results = response.get("results", [])

    for result in results:
        text = f"Title: {result.get('title', '')}\nContent: {result.get('content', '')}\nURL: {result.get('url', '')}"

        metadata = {
            "title": result.get("title", ""),
            "url": result.get("url", ""),
            "score": result.get("score", 0),
            "source_type": "tavily_search",
            "query": query,
            "timestamp": datetime.now().isoformat()
        }

        memory_manager.write_knowledge_base(text, metadata)

    return results
```

### Why this is important
This is a **write-back loop**:
- the tool fetches new external knowledge
- the results are stored into semantic memory
- later turns can retrieve them as part of knowledge base memory

### In your project
This is a very strong pattern.
Do not treat tool results as transient unless they truly are.
If the result is reusable, persist it.

Examples to persist:
- web search hits
- fetched docs
- API reference lookups
- benchmark outputs
- design decisions
- tool execution summaries

---

## 9.4 Local tool example

```python
from datetime import datetime

@toolbox.register_tool(augment=True)
def get_current_time(detailed: bool = False) -> str:
    if detailed:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    else:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
```

### Why it matters
Not every tool should be an external API.
Local deterministic tools are cheap, fast, and reliable.

### Use cases
- time/date
- string transforms
- file parsing
- validation
- unit conversion
- basic math
- feature extraction

---

## 9.5 arXiv candidate search tool

```python
import json
from urllib.parse import urlparse

def _arxiv_id_from_entry_id(entry_id: str) -> str:
    if not entry_id:
        return ""
    path = urlparse(entry_id).path
    return path.split("/abs/")[-1].strip("/")

@toolbox.register_tool(augment=False)
def arxiv_search_candidates(query: str, k: int = 5) -> str:
    docs = arxiv_retriever.invoke(query)
    candidates = []
    for d in (docs or [])[:k]:
        meta = d.metadata or {}
        entry_id = meta.get("Entry ID", "")
        candidates.append({
            "arxiv_id": _arxiv_id_from_entry_id(entry_id),
            "entry_id": entry_id,
            "title": meta.get("Title", ""),
            "authors": meta.get("Authors", ""),
            "published": str(meta.get("Published", "")),
            "abstract": (d.page_content or "")[:2500],
        })
    return json.dumps(candidates, ensure_ascii=False, indent=2)
```

### Why this is used
This is a discovery tool rather than a direct ingestion tool.
It returns structured candidate papers with metadata, which lets the agent:
- inspect candidate options
- choose a relevant paper
- then fetch the full document only when needed

### Design lesson
Separate:
- **search/discovery**
from
- **heavy ingestion/fetch**

This reduces unnecessary cost and noise.

---

## 9.6 Deep ingestion tool for papers

```python
from datetime import timezone
from langchain_community.document_loaders import ArxivLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter

@toolbox.register_tool(augment=True)
def fetch_and_save_paper_to_kb_db(
    arxiv_id: str,
    chunk_size: int = 1500,
    chunk_overlap: int = 200,
) -> str:
    loader = ArxivLoader(
        query=arxiv_id,
        load_max_docs=1,
        doc_content_chars_max=None,
    )
    docs = loader.load()
    if not docs:
        return f"No documents found for arXiv id: {arxiv_id}"

    doc = docs[0]
    title = doc.metadata.get("Title") or doc.metadata.get("title") or f"arXiv {arxiv_id}"
    entry_id = doc.metadata.get("Entry ID") or doc.metadata.get("entry_id") or ""
    published = doc.metadata.get("Published") or doc.metadata.get("published") or ""
    authors = doc.metadata.get("Authors") or doc.metadata.get("authors") or ""

    full_text = doc.page_content or ""
    if not full_text.strip():
        return f"Loaded arXiv {arxiv_id} but extracted empty text (PDF parsing issue)."

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
    )

    chunks = splitter.split_text(full_text)

    ts_utc = datetime.now(timezone.utc).isoformat()
    metadatas = []
    for i in range(len(chunks)):
        metadatas.append(
            {
                "source": "arxiv",
                "arxiv_id": arxiv_id,
                "title": title,
                "entry_id": entry_id,
                "published": str(published),
                "authors": str(authors),
                "chunk_id": i,
                "num_chunks": len(chunks),
                "ingested_ts_utc": ts_utc,
            }
        )

    memory_manager.write_knowledge_base(chunks, metadatas)

    return (
        f"Saved arXiv {arxiv_id} to knowledge base: "
        f"{len(chunks)} chunks (title: {title})."
    )
```

### Why this is important
This is the full ingestion pattern:
1. fetch long document
2. normalize metadata
3. chunk the content
4. embed/store chunks
5. keep metadata for later retrieval

### In your project
Use this exact pattern for:
- PDFs
- research papers
- product manuals
- meeting transcripts
- design docs
- source repositories

Chunk metadata is critical because later you may want:
- attribution
- source reconstruction
- chunk ordering
- selective filtering

---

## 10. L4 Notebook: Memory Operations, Summarization, Consolidation

`L4.ipynb` is about memory lifecycle operations, especially when context starts to overflow.

---

## 10.1 Context engineering

Context engineering means curating which information should go into the model context window.
The principle is not "put everything in the prompt."
It is "maximize signal per token."

In practice:
- recent conversation is useful
- relevant KB passages are useful
- prior workflow patterns may be useful
- entities may disambiguate references
- older conversation should usually be compressed
- irrelevant memory should stay out

---

## 10.2 Context usage estimator

```python
def calculate_context_usage(context: str, model: str = "gpt-5-mini") -> dict:
    estimated_tokens = len(context) // 4
    max_tokens = MODEL_TOKEN_LIMITS.get(model, 128000)
    percentage = (estimated_tokens / max_tokens) * 100
    return {"tokens": estimated_tokens, "max": max_tokens, "percent": round(percentage, 1)}
```

### Why this is used
This is a simple heuristic to monitor context occupancy before making another model call.

### In your project
Even if you later replace this with a tokenizer-specific estimate, the control logic is correct:
- below threshold -> continue
- near threshold -> warn
- above threshold -> summarize or compact

---

## 10.3 Context status monitor

```python
def monitor_context_window(context: str, model: str = "gpt-5-mini") -> dict:
    result = calculate_context_usage(context, model)

    if result['percent'] < 50:
        result['status'] = 'ok'
    elif result['percent'] < 80:
        result['status'] = 'warning'
    else:
        result['status'] = 'critical'

    return result
```

### Why this is used
This creates simple, deterministic thresholds for memory management decisions.

### In your project
Recommended extension:
- `<50%`: do nothing
- `50–80%`: prefer retrieval tightening, fewer chunks, fewer tools
- `>80%`: summarize/offload old conversation
- `>90%`: force compaction and reduce tool result verbosity

---

## 10.4 Summarize arbitrary context and store it

```python
def summarise_context_window(content: str, memory_manager, llm_client, model: str = "gpt-5-mini") -> dict:
    cleaned = (content or "").strip()
    if not cleaned:
        return {"status": "nothing_to_summarize"}

    summary_prompt = f"""You are creating durable memory for an AI research assistant.
Summarize this conversation so it can be resumed accurately later.

Output with exactly these headings:
### Technical Information
### Emotional Context
### Entities & References
### Action Items & Decisions
...
Conversation:
{cleaned[:6000]}"""

    response = llm_client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": summary_prompt}],
        max_completion_tokens=4000
    )
    ...
    description = ...
    summary_id = str(uuid.uuid4())[:8]
    memory_manager.write_summary(summary_id, cleaned, summary, description)

    return {"id": summary_id, "description": description, "summary": summary}
```

### Why this is good design
This summary is:
- durable
- structured
- optimized for continuation
- labeled with a short human/agent-friendly description
- stored in summary memory for later retrieval

### Why the headings matter
The headings separate different forms of continuation state:
- **Technical Information** → facts, APIs, decisions, errors
- **Emotional Context** → useful in assistant continuity or support settings
- **Entities & References** → names and identifiers
- **Action Items & Decisions** → unresolved tasks and commitments

This is much better than a flat prose summary.

### In your project
Use a structured summary schema like:

```text
### Technical Information
### Constraints
### Entities & References
### Decisions
### Open Questions
### Next Actions
```

Choose headings based on your domain.

---

## 10.5 Expand summary back into detailed memory

```python
@toolbox.register_tool(augment=True)
def expand_summary(summary_id: str) -> str:
    """
    Expand a summary reference to retrieve the original conversations.
    """
    summary_text = memory_manager.read_summary_memory(summary_id)
    original_conversations = memory_manager.read_conversations_by_summary_id(summary_id)

    return f"""
            ## Summary Context
                {summary_text}

                {original_conversations}
        """
```

### Why this matters
Summarization is lossy.
This tool restores detail when the summary is insufficient.

### Example
Summary memory says:
- "[Summary ID: 8a2b19cd] MemGPT paper retrieval and key insight summary"

Later the user asks:
- "What exactly was my first question?"

The summary alone may be insufficient.
The agent can call `expand_summary("8a2b19cd")` and inspect the original messages.

This is one of the strongest ideas in the course:
**summarize to save context, but preserve recoverability via database-backed expansion.**

---

## 10.6 Summarize exact unsummarized conversation units

```python
def summarize_conversation(thread_id: str) -> dict:
    with memory_manager.conn.cursor() as cur:
        cur.execute(f"""
            SELECT id, role, content, timestamp
            FROM {memory_manager.conversation_table}
            WHERE thread_id = :thread_id AND summary_id IS NULL
            ORDER BY timestamp ASC
        """, {"thread_id": thread_id})
        rows = cur.fetchall()

    if not rows:
        return {"status": "nothing_to_summarize"}

    message_ids = []
    transcript_lines = []
    for msg_id, role, content, timestamp in rows:
        message_ids.append(msg_id)
        ts_str = timestamp.strftime('%Y-%m-%d %H:%M:%S') if timestamp else "Unknown"
        transcript_lines.append(f"[{ts_str}] [{str(role).upper()}] {content}")

    transcript = "\n".join(transcript_lines)

    result = summarise_context_window(transcript, memory_manager, client)
    if result.get("status") == "nothing_to_summarize":
        return result

    summary_id = result["id"]

    with memory_manager.conn.cursor() as cur:
        cur.executemany(f"""
            UPDATE {memory_manager.conversation_table}
            SET summary_id = :summary_id
            WHERE id = :id AND summary_id IS NULL
        """, [{"summary_id": summary_id, "id": msg_id} for msg_id in message_ids])
    memory_manager.conn.commit()

    result["num_messages_summarized"] = len(message_ids)
    return result
```

### Why this is a very important implementation pattern
This function does **precise source accounting**:
- summarizes only rows not already summarized
- writes a new summary unit
- marks exactly those rows with the summary ID

This prevents:
- duplicate summarization
- summary drift
- inability to trace which messages were compacted

### In your project
This pattern is strongly recommended.
Do not summarize conversation in an ad hoc way without provenance.

---

## 10.7 Context compaction with section replacement

```python
def offload_to_summary(context: str, memory_manager, llm_client, thread_id: str = None) -> tuple:
    if thread_id:
        result = summarize_conversation(thread_id)
    else:
        result = summarise_context_window(raw_context, memory_manager, llm_client)

    if result.get("status") == "nothing_to_summarize":
        return raw_context, []

    summary_ref = f"[Summary ID: {result['id']}] {result['description']}"
    conversation_stub = (
        "## Conversation Memory\n"
        "Older conversation content was summarized to reduce context size.\n"
        "Use Summary Memory references + expand_summary(id) for full detail."
    )
    ...
```

### Why this is used
This function **does not destroy the whole context**.
It specifically replaces only the conversation-heavy section with a stub and injects a summary reference.

That is the right design:
- preserve knowledge memory
- preserve entity memory
- preserve workflow memory
- only compact the highest-volume, lowest-precision section

### In your project
Avoid global summarization of the entire prompt.
Prefer selective section compaction.

---

## 10.8 Agent-accessible summary tool

```python
@toolbox.register_tool(augment=True)
def summarize_and_store(text: str, thread_id: str = None) -> str:
    if thread_id:
        result = summarize_conversation(thread_id)
        if result.get("status") == "nothing_to_summarize":
            return f"No unsummarized messages found for thread {thread_id}."
        return f"Stored as [Summary ID: {result['id']}] {result['description']}"

    result = summarise_context_window(text, memory_manager, client)
    if result.get("status") == "nothing_to_summarize":
        return "No content to summarize."
    return f"Stored as [Summary ID: {result['id']}] {result['description']}"
```

### Why this matters
The agent itself can now decide to summarize and persist memory when needed.
That is part of becoming memory-aware.

---

## 11. L5 Notebook: Full Memory-Aware Agent

`L5.ipynb` combines all earlier patterns into one memory-aware agent loop.

---

## 11.1 System prompt teaches memory semantics

```python
AGENT_SYSTEM_PROMPT = """
# Role
You are a memory-aware agentic research assistant with access to tools.

# Context Window Structure (Partitioned Segments)
The user input is a partitioned context window. It contains a `# Question` section followed by memory segments.
Treat each segment as a distinct memory store with a specific purpose:
- `## Conversation Memory`
- `## Knowledge Base Memory`
- `## Workflow Memory`
- `## Entity Memory`
- `## Summary Memory`

# Memory Store Semantics
- Conversation Memory: Recent thread-level dialogue and instructions. Use it for continuity, user preferences, and unresolved requests.
- Knowledge Base Memory: Retrieved documents/passages. Use it to ground factual and technical claims.
- Workflow Memory: Prior execution patterns and step sequences. Use it to plan tool usage; adapt patterns, do not copy blindly.
- Entity Memory: Named people/orgs/systems and descriptors. Use it to disambiguate references and keep naming consistent.
- Summary Memory: Compressed older context represented by summary IDs. When thread-scoped summaries exist, prefer summaries for the active thread_id.

# Summary Expansion Policy
If critical detail is only present in Summary Memory or appears ambiguous, call `expand_summary(summary_id)` before relying on it.
...
"""
```

### Why this is critical
The model is explicitly told:
- what each memory section means
- how to prioritize them
- when to expand summaries
- how to treat conflicts

This is the difference between:
- "memory exists"
and
- "the model knows how to use memory correctly"

### In your project
If you have multiple memory streams, the agent prompt should define:
- section names
- semantics
- retrieval policy
- conflict resolution order
- when to call memory tools

---

## 11.2 Tool execution wrapper

```python
def execute_tool(tool_name: str, tool_args: dict, current_thread_id: str | None = None) -> str:
    if tool_name not in toolbox._tools_by_name:
        return f"Error: Tool '{tool_name}' not found"

    args = dict(tool_args or {})

    if tool_name == "summarize_and_store" and "thread_id" not in args and current_thread_id is not None:
        args["thread_id"] = str(current_thread_id)

    return str(toolbox._tools_by_name[tool_name](**args) or "Done")
```

### Why this is used
This wrapper injects agent runtime context into tool calls.
For summarization, it ensures the active thread ID is passed so exact conversation rows can be marked summarized.

### In your project
Your tool wrapper is the right place to enforce:
- thread scoping
- security rules
- default arguments
- context-aware parameter injection
- logging hooks

---

## 11.3 Main memory-aware agent loop

```python
def call_agent(query: str, thread_id: str = "1", max_iterations: int = 10) -> str:
    thread_id = str(thread_id)
    steps = []
    summaries = []

    memory_context = ""
    memory_context += memory_manager.read_conversational_memory(thread_id) + "\n\n"
    memory_context += memory_manager.read_knowledge_base(query) + "\n\n"
    memory_context += memory_manager.read_workflow(query) + "\n\n"
    memory_context += memory_manager.read_entity(query) + "\n\n"
    memory_context += memory_manager.read_summary_context(query, thread_id=thread_id) + "\n\n"

    usage = calculate_context_usage(memory_context)
    if usage['percent'] > 80:
        memory_context, summaries = offload_to_summary(
            memory_context,
            memory_manager,
            client,
            thread_id=thread_id,
        )

    context = f"# Question\n{query}\n\n{memory_context}"

    dynamic_tools = memory_manager.read_toolbox(query, k=5)

    memory_manager.write_conversational_memory(query, "user", thread_id)
    try:
        memory_manager.write_entity("", "", "", llm_client=client, text=query)
    except Exception:
        pass

    messages = [
        {"role": "system", "content": AGENT_SYSTEM_PROMPT},
        {"role": "user", "content": context}
    ]
    ...
```

### Why this loop is good
It does the right things in the right order:

1. retrieve context from different stores
2. measure context size
3. compact if needed
4. prepend the current question
5. retrieve only relevant tools
6. write the new user message into conversational memory
7. extract entities from the query
8. run tool-calling loop
9. log tool outputs
10. write workflow memory
11. write entities from final answer
12. write assistant response to conversation memory

That is a genuine memory lifecycle.

---

## 11.4 Tool-call persistence and bounded context control

Inside the loop, tool outputs are persisted in tool log memory.

```python
log_id = memory_manager.write_tool_log(
    thread_id=thread_id,
    tool_call_id=tc.id,
    tool_name=tool_name,
    tool_args=tool_args,
    result=result,
    status=status,
    error_message=error_message,
    metadata={"iteration": iteration + 1},
)

if len(result) > 3000:
    result_for_llm = result[:3000] + f"\n\n[Truncated for context. Full output saved in TOOL_LOG_MEMORY as log_id: {log_id}]"
else:
    result_for_llm = result
```

### Why this matters
This is another excellent pattern:
- keep the full result in persistent memory
- only pass a bounded version back to the LLM context
- preserve traceability and reproducibility

### In your project
Use this for:
- long API responses
- scraped web pages
- verbose logs
- database query results
- large code analysis outputs

Persist full output, send the model only what it can use.

---

## 11.5 Workflow write-back

At the end of the loop:

```python
if steps:
    memory_manager.write_workflow(query, steps, final_answer)
```

### Why this is important
Each completed task becomes future procedural memory.

### Example
The agent solves:
- query: "Find the MemGPT paper and summarize its core idea"

Workflow memory may store:
- searched toolbox for arXiv-related tools
- retrieved arXiv candidates
- fetched full paper
- stored chunks in KB
- answered from KB

Later, a similar research query can retrieve that workflow as a planning hint.

---

## 12. Context Window Strategy Recommended for Your Project

Since you said this KB will be used to update your context window and memory of your own agent, the most useful practical pattern is this:

### 12.1 Partitioned prompt layout

```text
# Question
{current_user_query}

## Conversation Memory
{recent_thread_messages}

## Knowledge Base Memory
{retrieved_passages}

## Workflow Memory
{similar_prior_task_patterns}

## Entity Memory
{relevant entities and descriptors}

## Summary Memory
{summary ids + descriptions for older context}
```

### 12.2 Retrieval policy
For every user turn:

1. **Conversation memory**
   - retrieve recent messages from the same thread
   - prefer unsummarized recent messages
   - optionally include the latest summary reference

2. **Knowledge base memory**
   - semantic search using the current query
   - optionally hybrid retrieval (keyword + vector)
   - rerank if available

3. **Workflow memory**
   - semantic search by current task
   - retrieve prior action sequences

4. **Entity memory**
   - extract key entities from current query
   - retrieve matching entity units

5. **Summary memory**
   - retrieve thread-scoped summaries first
   - only expand if needed

6. **Toolbox memory**
   - retrieve top-k tools from toolbox store
   - do not pass entire tool catalog

### 12.3 Compaction policy
- if context usage > 80%, summarize older conversation
- replace old raw conversation with a stub and summary ID
- preserve KB and workflow sections
- allow `expand_summary` to recover details

---

## 13. Recommended MemoryManager Interface for Your Project

If you are implementing this yourself, use a manager interface like this:

```python
class MemoryManager:
    # conversation
    def write_conversational_memory(self, text: str, role: str, thread_id: str, metadata: dict | None = None): ...
    def read_conversational_memory(self, thread_id: str, limit: int = 20) -> str: ...
    def read_conversations_by_summary_id(self, summary_id: str) -> str: ...

    # knowledge base
    def write_knowledge_base(self, text, metadata): ...
    def read_knowledge_base(self, query: str, k: int = 5) -> str: ...

    # workflow
    def write_workflow(self, query: str, steps: list[str], final_answer: str): ...
    def read_workflow(self, query: str, k: int = 3) -> str: ...

    # toolbox
    def register_tool(self, fn, augment: bool = False): ...
    def read_toolbox(self, query: str, k: int = 5): ...

    # entity
    def write_entity(self, subject, entity_type, description, llm_client=None, text: str | None = None): ...
    def read_entity(self, query: str, k: int = 5) -> str: ...

    # summary
    def write_summary(self, summary_id: str, original_text: str, summary: str, description: str): ...
    def read_summary_memory(self, summary_id: str) -> str: ...
    def read_summary_context(self, query: str, thread_id: str | None = None, k: int = 5) -> str: ...

    # logs
    def write_tool_log(self, thread_id: str, tool_call_id: str, tool_name: str, tool_args: dict,
                       result: str, status: str, error_message: str | None, metadata: dict | None = None): ...
```

---

## 14. Minimal Portable Version for Non-Oracle Stacks

The course uses Oracle AI DB, but the architecture is portable.

You can reproduce it with:
- PostgreSQL + pgvector
- SQLite + a vector sidecar
- Qdrant + Postgres
- Weaviate + SQL metadata
- Milvus + relational metadata
- LanceDB + SQL app state

### Portable schema idea
- `conversation_memory` → SQL
- `tool_log_memory` → SQL
- `knowledge_memory` → vector collection + metadata
- `workflow_memory` → vector collection + metadata
- `toolbox_memory` → vector collection + metadata
- `entity_memory` → vector collection + metadata
- `summary_memory` → vector collection + metadata

The important thing is not Oracle.
It is:
- separated memory types
- explicit read/write APIs
- compaction and expansion
- workflow persistence
- semantic tool retrieval

---

## 15. Recommended Enhancements Beyond the Course

If you are adapting this architecture into a real project, the following upgrades are worth adding.

### 15.1 Hybrid retrieval
The transcript mentions lexical, vector, graph, and hybrid retrieval.
Use hybrid retrieval for KB whenever possible:
- BM25 / keyword
- vector similarity
- reranker

### 15.2 Memory freshness and decay
Not all memories should persist equally.
Add:
- recency boosts
- time decay
- confidence scores
- soft archival for stale summaries

### 15.3 Conflict resolution
When memory units conflict:
- prioritize current user turn
- then latest conversation memory
- then verified KB evidence
- then workflow memory
- then older summaries

This logic already appears in the `AGENT_SYSTEM_PROMPT`.

### 15.4 Tool result normalization
Before writing tool output to KB:
- strip boilerplate
- normalize dates
- add source metadata
- separate raw text from cleaned extracted facts

### 15.5 Entity linking
Make entity memory stronger by linking aliases:
- "OpenAI"
- "Open AI"
- "the company"
- "their API team"

### 15.6 Evaluation hooks
Track:
- retrieval precision
- tool selection precision
- summary faithfulness
- summary expansion success
- workflow reuse value

---

## 16. End-to-End Example Flow

This is the intended end-to-end behavior of the final system.

### User turn 1
"Find the MemGPT paper."

Agent:
- retrieves toolbox tools relevant to papers
- calls arXiv search candidate tool
- chooses relevant paper
- calls paper fetch and ingestion tool
- writes paper chunks into KB
- responds with result
- logs tool outputs
- writes workflow memory
- writes conversation memory

### User turn 2
"What are its main ideas?"

Agent:
- retrieves recent conversation from same thread
- KB retrieval matches MemGPT chunks
- may retrieve entity memory for paper title/authors
- answers using KB chunks
- stores conversation and workflow

### User turn 3
Conversation gets long.

Agent:
- monitors context usage
- summarizes older conversation
- writes summary memory
- marks summarized rows using summary_id
- keeps summary reference in Summary Memory section

### User turn 4
"What was my very first question?"

Agent:
- summary reference suggests old context was compacted
- calls `expand_summary(summary_id)`
- recovers original messages
- answers correctly

This is the full memory engineering loop.

---

## 17. Direct Implementation Blueprint for Your Own Agent

If you want to apply the notebook architecture directly, implement your runtime in this order.

### Step 1 — Define stores
- conversation table
- tool log table
- KB vector store
- workflow vector store
- toolbox vector store
- entity vector store
- summary vector store

### Step 2 — Build MemoryManager
Expose read/write operations for each store.

### Step 3 — Add ingestion tools
- web search → write to KB
- doc fetch → chunk and write to KB
- paper fetch → chunk and write to KB
- local utilities

### Step 4 — Add toolbox memory
Store tool schemas as retrievable semantic memory.

### Step 5 — Build context assembler
Create partitioned prompt sections:
- question
- conversation
- KB
- workflow
- entity
- summary

### Step 6 — Add context monitor
Estimate current prompt size.
Trigger summarization when needed.

### Step 7 — Add summary tools
- summarize_and_store
- expand_summary

### Step 8 — Build agent loop
- retrieve context
- retrieve tools
- write user message
- run tool loop
- log outputs
- write workflow
- write final message

### Step 9 — Add evaluation
Track failures:
- missing memory retrieval
- wrong summary expansion
- wrong tool selection
- poor workflow reuse

---

## 18. Key Design Principles to Preserve

1. **Do not treat all memory as chat history.**
2. **Separate memory stores by purpose.**
3. **Use SQL for chronological episodic memory.**
4. **Use vectors for semantic retrieval-heavy stores.**
5. **Persist reusable tool outputs.**
6. **Summarize selectively, not globally.**
7. **Keep summaries expandable.**
8. **Teach the model how to interpret memory sections.**
9. **Use workflow memory to improve future planning.**
10. **Use toolbox memory to scale tool count safely.**

---

## 19. Copy-Paste Prompt Template for a Memory-Aware Agent

```text
# Role
You are a memory-aware agent with access to tools and persistent external memory.

# Context Window Structure
The prompt contains partitioned memory segments. Use each for its intended purpose.

# Memory Types
- Conversation Memory: recent thread messages, user instructions, unresolved requests
- Knowledge Base Memory: factual grounding and retrieved domain passages
- Workflow Memory: prior successful step sequences for similar tasks
- Entity Memory: named entities, aliases, and reference anchors
- Summary Memory: compressed older context, expandable via summary IDs

# Rules
1. Use the current question as the highest priority.
2. Prefer latest thread conversation over older summaries.
3. Use KB for factual grounding.
4. Use workflow memory as a planning hint, not a script.
5. If a summary reference may contain critical detail, call expand_summary(summary_id).
6. Use the minimum necessary tool calls.
7. If a tool returns long output, rely on stored logs and avoid flooding context.
```

---

## 20. Copy-Paste Pseudocode for Your Project

```python
def handle_user_turn(query: str, thread_id: str):
    # 1. Retrieve memory
    conversation = mm.read_conversational_memory(thread_id)
    kb = mm.read_knowledge_base(query)
    workflow = mm.read_workflow(query)
    entities = mm.read_entity(query)
    summaries = mm.read_summary_context(query, thread_id=thread_id)

    # 2. Build prompt
    memory_context = f"""
# Question
{query}

## Conversation Memory
{conversation}

## Knowledge Base Memory
{kb}

## Workflow Memory
{workflow}

## Entity Memory
{entities}

## Summary Memory
{summaries}
""".strip()

    # 3. Compact if needed
    usage = monitor_context_window(memory_context)
    if usage["status"] == "critical":
        memory_context, _ = offload_to_summary(memory_context, mm, llm_client, thread_id=thread_id)

    # 4. Retrieve relevant tools
    tools = mm.read_toolbox(query, k=5)

    # 5. Persist user message
    mm.write_conversational_memory(query, role="user", thread_id=thread_id)
    mm.write_entity("", "", "", llm_client=llm_client, text=query)

    # 6. Run tool-calling loop
    answer, steps = run_agent_loop(memory_context, tools)

    # 7. Persist results
    if steps:
        mm.write_workflow(query, steps, answer)
    mm.write_entity("", "", "", llm_client=llm_client, text=answer)
    mm.write_conversational_memory(answer, role="assistant", thread_id=thread_id)

    return answer
```

---

## 21. Final Practical Recommendation

If this KB is going to update your own agent’s memory and context-window logic, the most important reusable patterns from the materials are:

- **partition your prompt by memory type**
- **store old conversation as summary references instead of raw history**
- **keep summary expansion available as a tool**
- **use toolbox memory instead of passing every tool every turn**
- **persist useful tool outputs into KB**
- **write workflow traces after successful runs**
- **extract and persist entities continuously**
- **let a memory manager own all storage logic**

That is the architecture the transcript and notebooks are actually teaching.

