# 🧪 YellowPad Coding Challenge: Smart Clause Insertion in Word Document

## 📌 Overview
This exercise simulates a core challenge that we use with our current MS Word Add-in: **intelligently inserting legal language into the correct place in a few Word documents, with the correct formatting**. Your goal is to create a React/Nextjs application and user interface to load a word document that can:
- Identify where a clause should go in each contract
- For each contract, insert the clause with the correct font, size, spacing, and formatting
- Handle edge cases like headings, numbering, and placement before/after other sections
- For clarity, write general code; that is, don't hard code to these examples

We can provide you with a video of our current Word Add-in tool, if you need it.  

You may use **any tools, libraries, or AI assistance** — speed and practicality are key (V0, Claude Code, Windsurf, Cursor, Gemini, Kiro, Copilot, etc).

## Questions
If anytihng is unclear, please feel free to ask via email or Slack, just as you would if you were working for a company.

## 🕑 Time Limit
Spend **no more than 2 hours** on this. We value quick iteration and resourcefulness over polish.  If you need to start and stop your work, that's fine. Just keep an work.md file in your project detailing what you did and how long it took.

## 🧩 The Challenge

### You are given:
- A few sample contracts in `.docx` format (provided in the repo)
- A few snippets of text (a new clause to insert)
- An few instructions like:

> _“Insert this clause as section 4.2, directly after the last paragraph in section 4.1. If a heading is needed, format it bold and underlined, and match the document’s style.”_

### Your task:
1. Parse the documents and find the correct insertion points based on the instruction
2. Insert the clauses with **correct placement and formatting**, matching each contract’s existing style
3. Return the updated `.docx` files that shows the result

## 💻 Tech Notes
- Use **TypeScript** with an open-source Word manipulation tool
- The clauses, instructions, and contracts are provided in this repo

## ✅ Submission Instructions
1. Fork this repo (or download the files)
2. Work locally
3. Record a short Loom video walking through:
   - Your approach
   - Challenges you encountered
   - Any AI tools you used
4. Email your Loom + code/repo link to: `scott@yellowpad.ai`

