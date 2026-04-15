I would like to know what this REPO is about.
I created it but could not get an understanding why?!
Puzzlled I am.

## Plan: Understanding the EmailProcessor Repository

**TL;DR** - This repository is a freight logistics email intelligence system that converts Outlook .msg files into structured markdown documentation with timeline summaries for FML Freight Solutions to manage equipment shipping quotes and communications.

**Steps**
1. Explore repository structure and key files (README.md, scripts in zScripts/, project folders)
2. Analyze email processing scripts to understand functionality (extraction, conversion, timeline generation)
3. Examine sample project folders to identify usage patterns and outputs
4. Infer business purpose from folder names, email contents, and generated markdown
5. Document findings with examples showing real-world application

**Relevant files**
- README.md — Basic overview and setup instructions
- email_processor-msg-perplex.py — Primary processing script using Perplexity API
- EmailProcessor.json — Catalog of processed projects
- Sample project: BA2861 — Example with multiple emails and timeline output

**Verification**
1. Run a sample script on test emails to confirm processing works
2. Check generated markdown files for accurate extraction of tables and timelines
3. Validate JSON catalog matches folder structure

**Decisions**
- Repository purpose: Automate conversion of freight logistics emails to searchable documentation
- Key functionality: Email metadata extraction, HTML table parsing, timeline generation
- Business context: FML Freight Solutions managing international equipment shipments

**Further Considerations**
1. Original creator confusion in README suggests the system evolved organically
2. Multiple AI API versions indicate experimentation with different models for optimal results

### Detailed Understanding Based on Exploration

**1. Repository Overview**  
This is an automated email processing system for freight logistics. It takes Outlook `.msg` email files, extracts their content (including embedded HTML tables with quotes/invoices), converts them to markdown format, and generates chronological timeline summaries. Used by FML Freight Solutions to organize communications about equipment shipments (mainly Caterpillar machinery) between South Africa, DRC, and other African destinations.

**2. Possible Outcome/Goal**  
The system was created to solve the problem of scattered email communications in logistics operations. Before this, quotes, invoices, and shipment details were buried in Outlook threads with no easy way to track timelines or extract key data. The goal is to create a searchable, structured knowledge base from email archives, enabling faster decision-making, better client communication, and automated extraction of critical logistics information like costs, deadlines, and action items.

**3. Examples to Show Why It Matters**  
Here are concrete examples from the repository that demonstrate the system's value:

These examples show how the system transforms unstructured email data into actionable business intelligence for freight operations.

### Human adding comments here ###


**Why am I doing this?**
The why, I get looped in on so much information of shipping stuff that it hurts my IT Brain.
So I created this vault (Obsidian speak) to condense down the noise and see facts against a pre-scribed list of tasks to be completed.
Before this, I would ask MS Copilot to summarise. This would take too much time. I spent more time asking to summarise than doing the work.
Also where I live, the internet is pretty poor, this is Africa..... and having a bunch of scripts run is way cleaner for me. End goal is to create a dashboard with pre-scheduled shippments in planning phase, live shipments, closed out shipments. So when I do my offline review, and co-ordinate some planning, I can focus on tasks for the next day.
The dashboard will be run locally, in a python based web service ro close to it.

Yes I am that lazy. But hey, fun project to keep the wider team updated on shippment standings.
This is linked to repo https://github.com/JUngererNZ/TRANSNET-VESSEL-UPDATES
