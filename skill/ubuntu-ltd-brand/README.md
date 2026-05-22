# Ubuntu LTD Brand Skill

A self-contained, portable AI skill that produces consistent on-brand documents for Ubuntu Land Trust and Developments and the Common Ground Initiative.

Put it into any AI tool. The AI will then write letters, emails, LOIs, decks, social posts, board memos, press releases, and pipeline write-ups in the right voice with the right facts and the right governance attribution.

---

## What's in this folder

```
ubuntu-ltd-brand/
├── SKILL.md         Master skill definition with Anthropic-format frontmatter
├── voice.md         Voice principles, do/don't word lists, vocabulary
├── facts.md         Canonical numbers, team bios, pipeline, governance
├── templates.md     Document scaffolds (letter, email, LOI, deck, etc.)
├── boilerplate.md   Approved org descriptions (12 / 30 / 100 / 250 words)
└── README.md        This file
```

Total size: ~30 KB. Lean enough for any AI's context window.

---

## Install — by tool

### Claude Code (CLI)

1. Copy the entire `ubuntu-ltd-brand/` folder into `~/.claude/skills/` on your machine
2. Restart Claude Code or run `/skills` to refresh
3. The skill auto-triggers when you mention Ubuntu LTD, Common Ground, or any of the trigger phrases listed in `SKILL.md`

```bash
# Quick install
git clone https://github.com/jasonallenjohn-cell/ubuntu-ltd-brand-guidelines.git
mkdir -p ~/.claude/skills/
cp -r ubuntu-ltd-brand-guidelines/skill/ubuntu-ltd-brand ~/.claude/skills/
```

### Claude Desktop / claude.ai

1. Open a new chat
2. Click the **Projects** sidebar → **Create a project** (or use an existing one)
3. Name it "Ubuntu LTD"
4. In the project's **Project knowledge**, upload all five files: `SKILL.md`, `voice.md`, `facts.md`, `templates.md`, `boilerplate.md`
5. In the project **system prompt** field, paste:

> You are working on materials for Ubuntu Land Trust and Developments (Ubuntu LTD) and the Common Ground Initiative. Whenever you draft any external-facing content, read all five project knowledge files first — SKILL.md, voice.md, facts.md, templates.md, boilerplate.md — and apply them as documented. Don't invent numbers, names, or programme mechanics; ask if you need a fact that isn't in facts.md.

All chats in this project now write on-brand by default.

### ChatGPT / GPT-5 / Custom GPTs

Create a Custom GPT:
1. Go to **Explore GPTs** → **Create**
2. **Name:** Ubuntu LTD Brand Writer
3. **Instructions:** paste the full content of `SKILL.md` plus an instruction to follow the uploaded files
4. **Knowledge:** upload all five files
5. **Capabilities:** keep Code Interpreter enabled if you want it to generate docs/decks

### Gemini / Google AI Studio

Create a System Instruction:
1. Concatenate all five files into a single text block (`SKILL.md` first, then voice, facts, templates, boilerplate)
2. Paste into the **System instructions** field
3. Or upload as a "Source" if using NotebookLM

### Other AI tools

Paste `SKILL.md` as a system prompt. When you ask the AI to draft something, also paste the relevant section from `voice.md`, `facts.md`, or `templates.md` based on what you need.

---

## How to use it (once installed)

Just ask normally. Examples:

> "Draft a letter to Mr. Mike Leone declining his contribution proposal politely."
>
> "Write a 250-word press release announcing the Launchpad ground-break in Barrie."
>
> "I need an email to the Network for the Advancement of Black Communities thanking Amanuel for the introduction to a potential operator."
>
> "Draft a one-paragraph LinkedIn post about hitting 5,623 units in pipeline."
>
> "What's our approved 100-word boilerplate?"
>
> "Write deck slides 1-5 for a CMHC briefing on the Common Ground Initiative."
>
> "Draft section 5 (Year-20 land purchase) of an LOI for 844 Veterans Drive."

The AI will pull the right voice, facts, and template — and produce something consistent with everything else Ubuntu LTD publishes.

---

## When to update this skill

- **Quarterly:** refresh `facts.md` with the current pipeline numbers (units, sites, Launchpad status)
- **When a deal closes:** update the Launchpad cohort table in `facts.md`
- **When the team changes:** update the bios in `facts.md`
- **When a partnership lands:** add to the partners list in `facts.md`
- **When voice norms evolve:** revise `voice.md` and update the do/don't list

To update: edit the file, commit to the repo (`github.com/jasonallenjohn-cell/ubuntu-ltd-brand-guidelines`), and notify the team. People who installed the skill via `git clone` can `git pull` to refresh.

---

## Source and updates

This skill lives in the Ubuntu LTD brand guidelines repository:
**https://github.com/jasonallenjohn-cell/ubuntu-ltd-brand-guidelines**

Folder path inside the repo: `/skill/ubuntu-ltd-brand/`

Direct downloads of each file: see the `/downloads/` page on the live site at **https://jasonallenjohn-cell.github.io/ubuntu-ltd-brand-guidelines/site/downloads.html**

---

## Why this exists

A new institution becomes recognizable through consistency. Same voice, same vocabulary, same numbers, same governance attribution, same insistence on the four-step mechanic — across every artifact the organisation produces. The brand isn't decoration; it's discipline.

This skill encodes that discipline so anyone on the team — or any partner drafting on Ubuntu LTD's behalf — produces work that fits the system, without needing to memorize 40 pages of guidelines.
