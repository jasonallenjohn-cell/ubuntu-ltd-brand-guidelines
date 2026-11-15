---
name: trillium-cobrand
description: Co-brand Trillium client material (deal deck, owner one-pager) with a realtor's brand so the realtor can present it to their clients. Trillium leads the layout and numbers; the realtor appears as presenter (cover band, header lock-up, closing contact slide, one-pager strip and footer). Strips internal lines (commission, bonus, refinance surplus, fees, profit) and blocks the build if any remain. TRIGGERS — "co-brand", "cobrand", "realtor version", "brand it for <realtor>", "presented by", "make a version Colin can show", a realtor's design system or brand guide shared with a Trillium deal.
---

# Trillium co-brand template

Trillium stays the main brand: green and gold layout, Archivo, the render, plans and numbers. The realtor is the
presenter, in their own dark colour, one accent and their typeface. One realtor profile drives every deal.

## Files
- `profiles/<realtor>.json`: one per realtor (start from `profiles/_template.json`). Keep real profiles and deal
  rewrites OUT of this repo: the Pages workflow publishes the whole repo root to the public site. Store them with the
  deal files and pass their paths to the script.
- `scripts/cobrand.py`: `deck`, `onepager`, `check` (run with no arguments for usage).
- `examples/example-rewrites.json`: the shape of a deal's client rewrites (stat swap + speaker notes).

## Steps
1. **Profile.** From the realtor's brand guide or design system, fill: name, brand, brokerage, brokerage_legal,
   registration, email, phone, social, website, colors (dark, accent, text, soft, muted, line), font (a Google
   Fonts family + css2 spec), logo_blob (after uploading a logo to the deck artifact). Use exact values from their
   files; never invent contact details. Put anything unverified in `to_confirm`.
2. `python3 scripts/cobrand.py check profiles/<realtor>.json`: fix or report every warning (placeholder 555-01xx
   phone, missing brokerage legal name / registration line for RECO).
3. **Deck.** `python3 scripts/cobrand.py deck <trillium deck root> profiles/<realtor>.json <out root>
   [--drop stress,next,refi] [--rewrites <deal>-rewrites.json]`. The source is a Trillium Slides deck folder
   (`project/deck.json` + `project/slides/`). Internal table rows are removed automatically; any other internal
   wording stops the build and names the slide: add a rewrite (client-safe stat or speaker notes) or drop the slide.
   Publish the output as a NEW Slides artifact (create from the Slides type, copy the image assets listed in
   `blobs_to_copy` from the source deck with `from_url` + `asset_ids`, swap the `/_blob/` ids, then publish
   `project/deck.json` + `deck.css` + cover, then the other slides). Never overwrite the internal Trillium deck.
4. **One-pager.** `python3 scripts/cobrand.py onepager "<Deal> 6 One Pager <Owner>.html" profiles/<realtor>.json
   <out>.html`, then render with the trillium-tpi-tpo-v5 `render_pdf.js` and look at it: one page, nothing hidden
   under the footer (the dark footer band always measures 1056 px; check the image, not the number).
5. Report: links, what was removed for clients, and the `check` warnings the realtor must confirm.

## Rules
- Only client-facing pieces are co-branded: deck, owner one-pager, Pro Forma Sample, design concept. Never the
  Partner Pro Forma, Master Model, Fee Summary, owner memo or the realtor's own commission summary.
- Every co-branded piece carries: "<Name> introduces property owners to Trillium; the TPI and TPO programs are
  offered and administered by Trillium only ... not an offer to sell securities." Have Trillium's lawyer approve
  the wording before first use with a new realtor.
- The realtor's accent is used once per area (presenter band, rules, labels); never mix their full palette into
  Trillium layouts.
