# Kinmo follow-up agent

![Tests](../../actions/workflows/tests.yml/badge.svg)

A small agent that keeps LinkedIn outreach for Kinmo Solutions moving. Every weekday morning it reads the outreach tracker, works out who is due a follow-up, and drafts a message for each person based on who they are and where the conversation stands. I review each draft and send it myself.

## The problem

Kinmo's clients come from LinkedIn conversations with importers, exporters, manufacturers and private equity firms. Most deals need two or three touches, but follow-ups slip when they depend on remembering who was contacted when. The tracker has the information. It just doesn't tell you what to do today.

## What it does

1. **Reads the tracker** (an Excel file) and keeps only LinkedIn contacts that are still open.
2. **Finds the last touch** from a `Last Contacted` column if there is one, or from dates in the comments, like `Proposal sent 22/10`. It also counts follow-ups already recorded.
3. **Decides who's due.** By default, that's no reply for 5 days and fewer than 2 follow-ups so far. Closed deals, people who replied, and referral contacts are left out.
4. **Drafts the message.** Each draft fits the client type (importer, exporter, manufacturer, PE firm, supplier) and the stage of the conversation (cold outreach, proposal sent, price objection, partnership, networking). The second follow-up is always a polite last note.
5. **Writes a review page** (`drafts/YYYY-MM-DD.html`). Each draft has a Copy button and a link to the person's LinkedIn profile, and the page also lists everyone else and why they aren't due.

| Stage | How it's detected | What the follow-up does |
|---|---|---|
| Proposal sent | "proposal" or "SOW" in comments | Checks whether they've reviewed it, and offers a walkthrough |
| Price objection | "expensive", "budget" or "cost" in comments | Offers a smaller first phase that fits their type of business |
| Sales | Purpose = Sales | Short value statement for their type, and asks for a 20-minute call |
| Partnership | Purpose = Partnership | Suggests referrals or joint projects |
| Network | Purpose = Network | Light check-in |

## Why it doesn't send messages itself

LinkedIn doesn't allow automated messaging, and accounts get restricted for it. The agent does the remembering and the writing. Sending stays a human step, which also means every message gets a final read.

## Set up

```bash
git clone https://github.com/OdunolaA/kinmo-followup-agent.git
cd kinmo-followup-agent
python3 -m pip install --user -r requirements.txt
cp config.example.json config.json      # then set tracker_path and sender_name
python3 followup.py --open
```

Try it without your own data: set `"tracker_path": "sample/network-tracker-sample.xlsx"` in `config.json`.

## Daily use

```bash
python3 followup.py --open              # draft today's follow-ups and open the page
python3 followup.py sent "Amanda Etuk"  # after sending, record it (resets the 5-day clock)
python3 followup.py list                # everyone in the tracker and their status
```

Run it every weekday at 9:00 on a Mac, with a notification:

```bash
bash scripts/schedule-mac.sh            # on
bash scripts/schedule-mac.sh --off      # off
```

If the scheduled run can't read the tracker, macOS is blocking background access to Documents. Go to **System Settings → Privacy & Security → Full Disk Access**, add `python3`, and the next run will work.

## Tracker columns

The agent reads the tracker's existing columns: `Client`, `Type`, `Contact Name`, `Contact Role`, `Purpose`, `Contact Method`, `Contract`, `Production` and `Comments`. Four optional columns make it more accurate:

| Column | Example | Used for |
|---|---|---|
| `Last Contacted` | 3/11/2026 | When you last messaged them (better than dates in comments) |
| `Follow-ups Sent` | 1 | Follow-ups sent before you started using the agent |
| `Replied` | Yes | Stops follow-ups once someone replies |
| `LinkedIn URL` | https://www.linkedin.com/in/... | Links straight to their profile instead of a search |

## Optional: personalise drafts with Claude

Set `"use_claude": true` in `config.json` and add an Anthropic API key with `export ANTHROPIC_API_KEY=...`. Each template draft is then rewritten using the tracker notes, kept under 90 words, with no invented facts. Without a key, the agent uses the templates in `templates.json`, which you can edit freely.

## Privacy

Your tracker, `config.json`, the follow-up log and the drafts all stay on your computer and are excluded by `.gitignore`. The repository only contains the code and a sample tracker with made-up contacts.

## Project structure

```
followup.py                  the agent: read tracker, decide who's due, draft, write page
templates.json               message templates by client type and stage
config.example.json          settings (copy to config.json)
sample/                      made-up tracker for trying it out
tests/                       unit tests, run on every push by GitHub Actions
scripts/schedule-mac.sh      weekday 9:00 schedule on macOS
```

## About

Built by Odun Akinmoladun, founder of Kinmo Solutions, with AI-assisted development (Claude).
