# VB6.tiny.clone – a tiny VB6-flavoured interpreter with a JSON + Qt GUI layer

> A very small, toy Visual Basic 6–like interpreter written in Python, with a JSON “runtime object model” and a PySide6 GUI bridge.  
> Designed as a fun experiment, not a production-grade language.

---

## What is this?

`VB6.tiny.clone` is a minimal scripting environment that **feels a bit like VB6**, but is implemented as:

- A tiny interpreter for a **VB-ish** subset (Subs, Functions, If/While/Do, simple expressions).
- A **JSON-based runtime object model** (for data and “types”).
- A small **PySide6 GUI layer** that lets scripts manipulate Qt widgets through VB-style code.

It’s intentionally small and hackable, meant more as a learning / nostalgia project than a serious language.

---

## Features

### Language

- `Sub` / `End Sub` and `Function` / `End Function` (currently parameterless).
- `If / ElseIf / Else / End If`
- `While / Wend`
- `Do / Loop`
- `Dim` and simple variable assignment.
- Basic expressions:
  - String and numeric literals.
  - Variables and dotted properties.
  - Function calls.
  - `&` for string concatenation.
  - Simple arithmetic (`+ - * /`) without full operator precedence.
- Simple comparisons in conditions (`=`, `<>`, `<`, `>`, `<=`, `>=`).

### JSON runtime

The interpreter exposes a mini “object system” based on JSON:

- `JsonNew(schemaName)` – create new JSON values based on a named schema.
- `JsonParse(text)` – parse JSON into a runtime object.
- `JsonStringify(value)` – convert runtime objects back to JSON text.
- `JsonGet(root, "path.to[0].value")` – read from nested JSON structures.
- `JsonSet(root, "path.to[0].value", newValue)` – write into nested JSON structures.

Schemas are defined in Python and give you something *like* VB `Type`s without extra syntax.

### GUI layer (PySide6)

There is also a small bridge from JSON to Qt widgets:

- **Forms** are defined in JSON.
- Widgets such as:
  - `TextBox`
  - `Label`
  - `Button`
  - `ListBox`
  - `ComboBox`
  - `WebBrowser` (QWebEngineView, if available)
- Simple data binding:
  - Each control can include a `bind` property pointing into a JSON data context.
- Event wiring:
  - Events are declared as:

    ```json
    "events": {
      "click": "btnSave_Click"
    }
    ```

  - The interpreter then looks up and runs `Sub btnSave_Click()` in the script.

---

## Demo

The main demo is `demo.zip`, which shows:

- A simple and minimalistic accounting app (demonstrating the data and working structure from the post found in: https://medium.com/@RobertKhou/double-entry-accounting-in-a-relational-database-2b7838a5d7f8). TODO: writting an app manual.


> **Note:** You need PySide6 (and Qt WebEngine for the browser control) installed for the demo to work.

---

## Installation

This isn’t published on PyPI (yet), so clone the repo and install dependencies manually:

```bash
git clone https://github.com/<your-username>/tinyvb.git
cd tinyvb

# Create & activate a virtualenv if you like, then:
pip install PySide6
# If your platform splits out WebEngine into a separate package, install that too.
```

Then:

```bash
python tinyvb_webext/demo2.py
```

If the web browser widget fails due to missing Qt WebEngine, the rest of the demo should still give you a feel for the interpreter and the JSON runtime.

---

## Project structure (suggested)

A possible layout for this repository:

```text
tinyvb/
├── tinyvb_webext/
│   ├── __init__.py
│   ├── runtime.py
│   ├── interpreter.py
│   ├── gui.py
│   └── demo2.py
├── examples/
│   └── mini.vb6.txt
├── docs/
│   └── design-notes.md
├── README.md
└── LICENSE
```

You can adjust this to taste, but this gives users a clear entry point and some extra documentation.

---

## AI-assistance disclosure

This project was created by a human author with significant assistance from
an AI coding assistant (OpenAI’s ChatGPT / GPT-5.1 Thinking).

The AI was used for:

- Brainstorming the architecture and feature set.
- Iterating on the interpreter logic and JSON runtime design.
- Helping refine the GUI bridge and example scripts.
- Producing and polishing documentation like this README.

All code and decisions were reviewed and accepted by a human before being committed.

---

## Limitations & non-goals

This is **not** a full VB6 re-implementation. Some limitations by design:

- No operator precedence beyond very simple expressions; use parentheses and simple binary operations.
- No error handling (`On Error` etc.).
- No parameterized procedures yet (all Subs/Functions are parameterless).
- Everything runs in a mostly **global** context.
- Only a handful of controls are implemented in the GUI layer.
- No security sandboxing – **do not** use this to run untrusted code.

Treat this as a toy interpreter and experiment.

---

## License

This software is released under [The Unlicense](https://unlicense.org/).

It is dedicated to the public domain. You can copy, modify, publish, use,
compile, sell, or distribute it for any purpose, commercial or
non-commercial, and by any means.
