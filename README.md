# Docx Engine for AI Agents

A Python library designed for precise .docx manipulation in LLM code execution environments.

This engine bridges the gap between AI Agents and Word documents. It allows agents to read document structure via JSON and write content (such as translations or edits) into specific locations without breaking the underlying XML structure.

## Key Features

* **Agent Native API:** The decoupled logic in `engine.py` returns clean JSON and Dictionaries, suitable for tool use and function calling.
* **Precise XML Injection:** Inserts paragraphs inside the document tree rather than just appending to the end, while preserving validity.
* **Chunking Support:** The `read_chunk` method allows processing large documents in small token windows.
* **Format Preservation:** Maintains existing styles, headers, and document hierarchy during edits.

## Quick Start

### 1. Installation

```bash
pip install python-docx lxml
```

### 2. Usage as a Library (AI Agent Workflow)
This is the core workflow for the Filesystem First approach where the agent manipulates the file directly in the container.

```python
from engine import DocxEngine

# 1. Load the document
doc = DocxEngine("contract.docx")

# 2. Get Machine Readable Structure
# Returns a JSON ready list of paragraph objects with IDs
structure = doc.get_structure_json()
print(structure)
# Output: [{"id": "p0", "text": "Service Agreement", "style": "Title"}, ...]

# 3. Apply Edits
# The agent targets specific IDs to insert content
doc.insert_translation(
    target_id="p0",
    translation_text="Dienstleistungsvertrag",
    style="Subtitle"
)

# 4. Save and Export
doc.save_file("contract_translated.docx")
```

### 3. Usage as CLI (Manual Testing)
The repository includes a REPL interface for manual testing and validation.

```bash
python main.py
```

```bash
> load test_doc.docx
> map
> insert_after p3 " [Reviewed by Claude]"
> save
```

## Technical Approach

I used `python-docx` to handle the file structure boilerplate but implemented custom logic for OXML manipulation to support mid-document insertion.

The standard library only allows appending paragraphs to the end of a document. To support insertion at specific indices, this engine accesses the underlying XML element (`._element`) of the target paragraph, locates its parent, and inserts the new node directly into the XML tree at the correct index.

```python
parent.insert(parent.index(target_xml) + 1, new_p_xml)
```

## Validation
The tool includes a validate command that attempts to unzip the output file and parse the internal document.xml using lxml. This catches XML corruption before the file is opened in Word.