---
name: tilda-article-layout
description: Верстка и визуальное оформление SEO-статей в Tilda через подключенный редактор: перенос структуры из Markdown, настройка блоков TL04/TL02A/TX01/ME606/FR310/TX16N2/QT07/TX20/T123, микроразметка, оглавление, таблицы, FAQ, чек-листы и точечная CSS-полировка. Использовать, когда нужно оформить готовую статью в открытой странице Tilda, а не просто написать текст.
metadata:
  short-description: Верстка SEO-статей в Tilda
---

# Tilda Article Layout

Use this skill when the user asks to place, arrange, or visually polish a prepared SEO article in an open Tilda page editor. The usual source is a Markdown article with SEO metadata and heading structure, and the target is an already-open authenticated Tilda editor controlled through Playwright/MCP.

## Core Rules

- Work against the currently open Tilda editor when available. Prefer Tilda's own editor functions over brittle clicks: `edrec__editRecordContent`, `edrec__sendForm('save','content')`, `tp__saveOnlyOneFieldInRecord`, `tp__dublicateRecord`, `tp__delRecord`, `tp__saveRecordsSort`, and `tp__updateRecord`.
- Do not publish the page unless the user explicitly asks. Say clearly when edits are saved but not published.
- Preserve existing user edits. Make narrowly scoped changes and avoid overwriting unrelated Tilda blocks or global CSS.
- Use stable record IDs and block codes from the live page. Verify after saving by reopening the form or re-reading the rendered record.
- Add all custom CSS in the existing `T123` code block using clear comment markers so each visual treatment can be replaced or removed independently.
- Avoid broad CSS selectors. Scope CSS to exact record IDs such as `#rec3123490101` to prevent changing header, footer, or unrelated article blocks.

## Article Structure

Use the Markdown article as the source of truth:

- `#` becomes the article H1/hero title.
- SEO `Title`, `Description`, URL slug, and metadata should be present in the prepared article file before layout work begins.
- `##` headings become Tilda `TL04` blocks.
- `###` headings become Tilda `TL02A` blocks unless the section is better represented by a specialized block.
- Body text under headings goes into `TX01` blocks.
- Keep the article order: heading, body/content block, related visual/link block, then the next heading.

Specialized sections:

- Use `ME606` for article navigation/table of contents. Fill it with main `H2` items and anchor links to the relevant `#rec...` records.
- Use `FR310` for checklist sections. Put each checklist item into a separate card and remove the old duplicate `TX01` checklist body.
- Use `TX16N2` for FAQ. Place all FAQ questions/answers there and remove duplicate `H3 + TX01` FAQ blocks.
- Use `QT07` for the conclusion/final takeaway when the user asks for a conclusion block.
- Use `TX20` for internal links to services/products. Distribute link blocks through the article near relevant sections rather than clustering all links in one place.
- Use `T123` for JSON-LD, custom scoped CSS, and small page-level code snippets.

## Tilda Workflow

1. Inspect the live page records under `#allrecords .record` and collect `recordid`, `data-record-cod`, and visible text.
2. Parse the Markdown article into ordered sections: H2, H3, body text, tables, checklists, FAQ, conclusion, CTA, links.
3. Reuse existing matching blocks first. Duplicate a nearby block only when more instances are needed.
4. Fill each block through the content form when possible:
   - Open: `await window.edrec__editRecordContent(recid)`.
   - Set native `input`/`textarea` values and dispatch `input` and `change` events.
   - Save: `window.edrec__sendForm('save', 'content')`.
   - Close: `window.edrec__closeEditForm(true)`.
5. For Zero Block text fields, use `tp__saveOnlyOneFieldInRecord(recordid, fieldName, fieldName, value)` when the ordinary content form fails.
6. Move records in the DOM to the correct order, then persist order with:
   ```js
   window.autosavesort_timer = setTimeout(() => {}, 60000);
   await window.tp__saveRecordsSort();
   ```
7. Verify the final structure by reading records around each changed section.

## Visual Treatment

Keep the visual language coherent across the article:

- Tables: convert plain/quill tables to styled responsive HTML tables inside `TX01`. Use a rounded white wrapper, dark header, subtle borders, soft shadow, warm accent column, and `overflow-x:auto` for mobile.
- Intro block: style the first `TX01` after `H2 Вступление` as an editorial lead. Make the first paragraph a lead card, convert the “В этой статье” label into a small dark pill, and style the list as compact cards.
- Checklist `FR310`: prefer numbered badges (`01`, `02`, ...) over large emoji icons. Keep the block background unchanged unless explicitly requested. Use scoped CSS for grid layout, card spacing, shadows, and mobile behavior.
- Conclusion `QT07`: if the user dislikes one visual direction, replace the previous CSS block rather than layering alternatives. Offer restrained variants: minimal editorial line, card, or quote-like final insight.
- Step headings: do not permanently style all H3 headings unless requested. If styling a sequence such as “Шаг 1...”, scope to exact record IDs and remove the CSS block cleanly if the user asks to roll it back.
- Image insertion: upload assets to Tilda CDN before referencing them. Do not point published content to local `D:\...` files. Be careful: `TX01` may sanitize `<img>` tags, leaving only captions. Prefer native image blocks, Zero Block image/shape fields, or a scoped CSS/CDN approach that is verified after save.

## SEO And Metadata

When preparing a Tilda article page, update both page settings and schema where applicable:

- Page URL alias should match the recommended slug, for example `blog/vnedrenie-sistemy-prodazh`.
- `T123` JSON-LD should use the final public URL, not the Tilda editor URL.
- Breadcrumb JSON-LD final item should match the article title and final URL.
- Article JSON-LD should include `mainEntityOfPage.@id`, `headline`, `datePublished`, author, publisher, and image when available.
- Use today's actual date when the user asks for today's publication date.
- Update visible breadcrumbs in `ME605` so the final breadcrumb item is the current article title.
- Update H1 and subtitle in the first block from the article title and description.

## Verification Checklist

Before finishing, verify the concrete outcomes relevant to the request:

- The requested blocks exist and contain the intended text.
- The visible order around edited sections is correct.
- Duplicates created during conversion were removed.
- All links point to intended anchors or service URLs.
- Custom CSS markers exist once and are scoped to exact records.
- Old CSS variants were removed when replacing a design.
- Tilda sanitization did not remove important HTML such as images or tables.
- The page is saved, and publication status is reported accurately.
