---
name: daily-report-form
description: Turn a free-text daily project status memo (a docx/text "справка") into two Word status reports for management — a detailed one and a short daily-dynamics one. Use when the user uploads or pastes a project status update (construction, contract execution, any multi-party work-in-progress report) and wants a recurring reporting form filled in from it, especially when they say this needs to happen daily/regularly.
---

# Daily report form

Converts an ad-hoc status memo into two standing report forms:

1. **Detailed form** — full breakdown by counterparty/contract: scope, schedule
   status, headcount/equipment plan vs. fact, documentation status, materials
   & supply, work closure (КС), advances & bank guarantees, a cross-cutting
   risk register, and a "suggestions to improve the form" section.
2. **Short/dynamic form** — a one-page dashboard: top KPIs, a per-party status
   table (RAG-coloured), a day-over-day "dynamics log" table meant to be
   extended with one new row per day, and an open-correspondence aging tracker.

This structure was designed against a real construction-project справка (a
generподряд with several subcontracts) but the renderer (`render.py`) is
domain-agnostic — it only renders whatever section/table structure you hand
it. Nothing about "contractors" or "construction" is hardcoded.

## When to use

- The user uploads/pastes a status memo and asks for a "форма отчётности" /
  daily report / status form to give to a manager, especially when they
  mention it recurs daily or want to track dynamics over time.
- The user asks to regenerate the same two forms from a *new* day's memo.

## Workflow

1. **Get the source text.** If it's a docx/pdf/etc., convert it with the
   `markitdown` CLI already in this repo: `python3 -m markitdown <file>`.
   If python-docx isn't installed yet for rendering (step 3), install it:
   `pip install --quiet python-docx` (safe to re-run).
2. **Extract facts, don't invent them.** Read the memo and fill the schema
   below per counterparty/contract. Where the memo is silent on something the
   form asks for (e.g. planned headcount, a КС submission date, a bank
   guarantee), use `gap()` — never fabricate a number or date. If the source
   itself contains internal contradictions (two different dates for the same
   contract, a rounding mismatch), keep both and flag them with a `top_notes`
   / section `notes` entry instead of silently picking one.
3. **Render.** Build two Python dicts (`detailed_spec`, `short_spec`) per the
   schema below, then:
   ```python
   import sys
   sys.path.insert(0, ".claude/skills/daily-report-form")
   from render import generate_detailed_report, generate_short_report, status, gap, blank, fmt

   generate_detailed_report(detailed_spec, "Справка_детальная_<объект>_<дата>.docx")
   generate_short_report(short_spec, "Справка_короткая_динамика_<объект>_<дата>.docx")
   ```
4. **Verify before sending.** Convert the generated docx back with
   `python3 -m markitdown <out>.docx` and read it — this catches malformed
   tables/typos cheaply, the same way you'd verify any generated artifact.
5. **Deliver.** Send both files to the user directly (e.g. via the
   host's file-delivery mechanism). **Do not commit filled-in reports that
   contain real business data into this repository** — the repo is shared
   Claude Code tooling, not a place for the user's confidential project data.
   Generated instances belong in a scratch/output directory outside git, or
   delivered straight to the user.
6. **Recurring runs.** For the next day's memo, repeat from step 1 — reuse
   the same `render.py` and the same `detailed_spec`/`short_spec` shape so
   the "Журнал динамики" table in the short form keeps accumulating
   comparable rows across days (carry forward prior rows if the user wants
   history retained, appending the new day).

## Helper functions (`render.py`)

- `fmt(n)` — RU-style thousands/decimal formatting: `1234567.8` → `"1 234 567,80"`.
- `status(text, level)` — coloured cell, `level` ∈ `"ok" | "warn" | "bad" | "note"`
  (green / amber / red / grey). Use for schedule status, RAG indicators.
- `gap(text="не указано в справке")` — grey italic cell marking a genuine gap
  in the source data.
- `blank(text="—")` — lighter grey italic cell for a future-date placeholder
  row (e.g. tomorrow's row in a dynamics log) — distinct from `gap()`, which
  means "the source should have had this and didn't."

## Schema

### `detailed_spec` (→ `generate_detailed_report`)

```python
{
  "title": str,                      # top H0 heading
  "subtitle_lines": [str, ...],      # bold lines under the title (object/parties/contract)
  "meta_line": str,                  # "Дата отчёта: ...    Составил: ___"
  "top_notes": [str, ...],           # ⚠ callouts right under the header (e.g. source data conflicts)
  "legend": bool,                    # render the ok/warn/bad/note colour legend
  "sections": [
    {
      "heading": str,                # e.g. "1. Наименование контрагента, виды работ"
      "intro": [str, ...],           # optional paragraphs before the table
      "table": {
        "headers": [str, ...],
        "rows": [[cell, ...], ...],  # cell = str | status(...) | gap(...) | blank(...)
        "widths_cm": [float, ...],   # optional, else auto
      },
      "notes": [str, ...],           # optional ⚠ callouts after the table
    },
    ...                              # one dict per report section — add as many as needed
  ],
  "suggestions": {                   # optional "how to improve this form" section
    "heading": str,
    "intro": str,
    "items": [{"lead": str, "text": str}, ...],  # lead is bolded, text is the explanation
  },
  "footer_note": str,                # optional italic note on its own page at the end
}
```

### `short_spec` (→ `generate_short_report`)

```python
{
  "title": str,
  "subtitle_lines": [str, ...],
  "meta_line": str,
  "blocks": [                         # same shape as `sections` above, no "intro" convention needed
    {"heading": str, "table": {...}, "notes": [str, ...]},
    ...                               # typically: KPI strip, per-party status, dynamics log, letters aging
  ],
  "footer_note": str,
}
```

## Recommended detailed-form section set

Derived from a real daily-reporting requirement; adjust per what the user
actually asks for, but this set has proven complete for a multi-contractor
project status report:

1. Наименование контрагента, виды проводимых работ по контракту
2. График производства работ (в графике/нет, отставание, догоняющие меры)
3. Численность людей и техники — план/факт
4. Наличие РД и сметной документации, выданной в производство
5. Наличие материалов/оборудования на площадке, договоры поставки, сроки, фронт работ
6. Закрытие работ — оперативное выполнение vs фактически принято, план по сдаче КС
7. Авансы и банковские гарантии
8. Проблемные вопросы и риски (cross-cutting, sourced from the memo's own "problem issues" section if present)
9. Предложения по форме справки

## Recommended short-form block set

1. Ключевые цифры (KPI strip — one-row table, 4-6 headline numbers)
2. Статус по контрагентам (one row per party: schedule status, closed sum/%, headcount, key blocker)
3. Журнал динамики (append-only day log: date, cumulative closed, %, headcount, overall schedule status, Δ vs. yesterday)
4. Открытые письма без ответа (aging tracker: letter #/date, addressee, subject, days outstanding)

## Gotchas

- Both report types render **landscape** by default (tables tend to be wide) — this is
  baked into `render.py`, no spec field controls it.
- `gap()` vs `blank()` matter: don't use `gap()` for a future date row in a
  dynamics log (that's not a data gap, it's an unfilled future) — use `blank()`.
- If the source memo has two different tables of quantities for the same
  contractor (e.g. cumulative-to-date vs. current-period), keep both as
  separate rows/columns rather than collapsing them — the distinction is
  usually meaningful (already-closed vs. submitted-but-not-yet-accepted).
- Recompute percentages/day-counts yourself from the raw dates/sums in the
  memo rather than trusting a percentage stated elsewhere in the source at
  face value — but if your computed figure and a stated one disagree, flag
  it, don't silently overwrite.
