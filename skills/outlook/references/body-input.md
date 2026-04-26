# Body & comment input

How to pass text content into the CLI's `--body` and `--comment` arguments cleanly. Applies to:

- `outlook mail draft --body ...`
- `outlook mail reply --body ...`
- `outlook mail forward --comment ...`
- `outlook calendar create --body ...`
- `outlook calendar update --body ...`
- `outlook calendar respond --comment ...`

All six accept the same input mechanics described below.

## Four ways to pass content — pick one

### (A) stdin / heredoc — recommended for agents

Cleanest for multi-line content. No escape quoting, no shell gymnastics:

```bash
outlook mail draft --to a@b.com --subject "Status" --body - <<'EOF'
Hi Alice,

Quick update on the deal:
  - All docs signed
  - Closing scheduled for Tuesday

Best,
Agent
EOF
```

```bash
outlook calendar create --subject "Project review" \
  --start 2026-05-15T10:00 --end 2026-05-15T11:00 \
  --attendees alice@example.com \
  --body - <<'EOF'
Agenda:
  1. Status update
  2. Risks + blockers
  3. Next steps

Pre-read in shared drive.
EOF
```

The `<<'EOF' ... EOF` heredoc preserves real newlines from the source; the single quotes around `'EOF'` mean no shell expansion, so what you write is what gets sent.

`--body -` (a single dash) tells the CLI to read the body from stdin. `--comment -` works the same way for `mail forward` and `calendar respond`.

### (B) `--body-file <path>` (mail only)

When the body already lives in a file:
```bash
outlook mail draft --to a@b.com --subject "..." --body-file /tmp/email.txt
```

No escape interpretation — the file is treated as raw UTF-8 text.

(Calendar commands don't have `--body-file`. Use stdin via heredoc instead.)

### (C) `--body` / `--comment` with `\n` escapes — the convenient one-liner

The CLI decodes `\n`, `\r`, `\t`, `\\` like `printf` does. So this works:

```bash
outlook mail draft --to a@b.com --subject "Hi" \
  --body "Hi Alice,\n\nQuick update:\n  - point 1\n  - point 2\n\nBest,\nAgent"
```

```bash
outlook calendar create --subject "Standup" \
  --start 2026-05-15T09:00 --end 2026-05-15T09:15 \
  --body "Agenda:\n  - Status\n  - Blockers\n  - Asks"
```

Pass `--raw-body` (or `--raw-comment` on forward / respond) to disable this decoding if you actually want a literal `\n` to appear in the output.

### (D) `--html` with `<br>` / `<p>` (mail only)

For properly-formatted HTML emails:
```bash
outlook mail draft --to a@b.com --subject "Update" --html \
  --body "<p>Hi Alice,</p><p>Quick update:</p><ul><li>point 1</li><li>point 2</li></ul><p>Best,<br>Agent</p>"
```

Use HTML when you need rich formatting (bold, lists, tables, links) that plain text can't carry.

(Calendar event bodies do NOT support HTML — Outlook renders calendar bodies as plain text only.)

## Anti-pattern — DO NOT do this

```bash
# WRONG: bash double-quotes do NOT interpret \n.
# Without the CLI's escape decoding (option C), this would email the literal
# four-character string "Hi\n\nbody" to the recipient.
# Current CLI auto-decodes, but option A (stdin/heredoc) is still preferred —
# it's the most readable and survives any future escape-handling change.
outlook mail draft --body "Hi\n\nbody"   # works now via auto-decode, but obscures intent
```

## What does NOT get escape-decoded

- `--subject` (subjects shouldn't have newlines anyway)
- `--location` (calendar)
- email addresses, recipient lists, etc.
- file contents (`--body-file`, stdin)

Only `--body` and `--comment` go through the decoder.

## Decision matrix

| Content shape | Use option |
| --- | --- |
| One line, no formatting | `--body "..."` (option C) |
| Multi-line, plain text, agent call | stdin / heredoc (option A) |
| Multi-line, plain text, content already in a file | `--body-file <path>` (option B, mail only) |
| Rich formatting (lists, bold, links) | `--html` + HTML body (option D, mail only) |
| Calendar event body with formatting | stdin/heredoc plain text (option A) — calendar bodies are plain text only |
