# 🧪 YellowPad Coding Challenge: Smart Clause Insertion in MS Word

## 📌 Overview
This exercise simulates a core challenge from our MS Word Add-in: **intelligently inserting legal language into the correct place in a few Word documents, with the correct formatting**. Your goal is to:

- Identify where a clause should go in each contract
- For each contract, insert the clause with the correct font, size, spacing, and formatting
- Handle edge cases like headings, numbering, and placement before/after other sections
- For clarity, write general code; that is, don't hard code to these examples

You may use **any tools, libraries, or AI assistance** — speed and practicality are key.

## 🕑 Time Limit
Spend **no more than 2 hours** on this. We value quick iteration and resourcefulness over polish.

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
- Use **Python** with `python-docx` (or similar), or **TypeScript** with an open-source Word manipulation tool
- You **do not need to use LLMs**, but feel free to if it helps
- The clauses, instructions, and contracts are provided in this repo

## ✅ Submission Instructions
1. Fork this repo (or download the files)
2. Work locally
3. Record a short Loom video walking through:
   - Your approach
   - Challenges you encountered
   - Any AI tools you used
4. Email your Loom + code/repo link to: `ananda@yellowpad.ai`
