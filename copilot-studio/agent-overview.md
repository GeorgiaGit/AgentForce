
<p align="center">
  <img src="https://github.com/GeorgiaGit/AgentForce/blob/main/assets/articleforge_small.png" alt="Description of Image"/>
</p>

# Copilot Studio Agent Overview

**Track**: Enterprise Agents (Microsoft Agents League)

**Built with**: Microsoft Copilot Studio + Azure Functions

**Agent**: **Article Forge** (KB Orchestrator)

**Submission type**: Documentation-based (no Copilot Studio solution export)

**Public-repo note**: This overview intentionally avoids tenant IDs, internal URLs, and proprietary system names. Names are generalized (e.g., “SharePoint KB site”) while preserving the real architecture and behavior.

---

## 1. Agent Purpose

**Agent Name**: **Article Forge — KB Orchestrator**

Article Forge is a **deterministic documentation generator** for IT knowledge base (KB) content. It transforms either:

- a **user-provided scenario** (typed prompt), or
- an **uploaded rough-draft PDF**

…into a **polished, publication-ready KB article**.

**Primary goals:**

- Produce consistent KB outputs in strictly enforced formats.
- Ground the resulting KB content using approved enterprise knowledge sources, Microsoft authoritative resources, and safe general technical knowledge.
- Output a canonical Markdown KB article (Adaptive Card-safe), plus end-user SharePoint page copy and a template-ready Word document version.
- Support downstream publication by providing structured content that can be consumed by an **Azure Function** to create a **SharePoint page**.

---

## 2. Target Users & Scenarios

**Intended users:**

- IT service desk / helpdesk teams
- IT documentation owners / knowledge managers
- MSP partners creating standardized KB content

**Primary scenarios:**

1. **Scenario-based KB generation**
   - User types a KB request (e.g., “Clear Teams cache on Windows 11 and macOS; include verification”).
   - Agent produces the three standardized outputs.

2. **Draft-to-KB transformation (PDF upload)**
   - User uploads a rough draft or existing doc as a PDF.
   - Agent extracts the intent and steps, reconciles gaps, and outputs a refined KB article in the required formats.

3. **Publish-ready SharePoint KB**
   - The Markdown article (Output 1) serves as the canonical review version.
   - The SharePoint page content (Output 2) is ready for automated posting via a publish workflow (Azure Function).

---

## 3. High-Level Architecture

The solution uses a **content-orchestration + secure publication** pattern:

- **Copilot Studio (Article Forge)**
  - Collects user input (scenario text or uploaded PDF)
  - Applies deterministic formatting rules and ambiguity handling
  - Produces KB outputs for downstream use

- **Knowledge Grounding (conceptual)**
  - Internal docs / approved KB sources (enterprise knowledge)
  - Microsoft authoritative documentation (product behavior, UI labels, version caveats)
  - Safe general technical knowledge for common troubleshooting patterns

- **Azure Functions (Publication boundary)**
  - Receives the generated SharePoint page content
  - Creates/updates a SharePoint page in the target KB site
  - Stores metadata/traceability (e.g., title, category, version) as needed

```text
User (scenario text or PDF)
  → Copilot Studio Agent: Article Forge (KB Orchestrator)
    → Grounding sources (enterprise + Microsoft authoritative)
    → Outputs (Markdown KB + SharePoint page text + Word template text)
      → Azure Function (publish)
        → SharePoint KB site (page created/updated)
```

**Diagram**: See `architecture/architecture-diagram.png` (or include the Mermaid diagram from this repo) for a one-page overview.

---

## 4. Agent Capabilities

### Input Modalities

- **Typed scenario**: user describes the KB topic and constraints
- **PDF upload**: user provides a rough draft; agent refines and standardizes

### Deterministic Output Generation

- Produces one or more outputs based on selection rules
- Default behavior is to generate **all three outputs** in a fixed order
- Uses strict headers and formatting guardrails to keep outputs machine-consumable

### Quality & Safety Guardrails

- Avoids secrets, tokens, and tenant-sensitive details
- Adds cautions for steps that may impact user data or cached credentials
- Includes Windows and macOS variants when applicable
- Calls out version/build/policy caveats when they could change steps

---

## 5. Topics & Intent Design

The agent is organized around deterministic routing and generation.

| Topic / Capability | Purpose | Notes |
|---|---|---|
| KB – Entry Router | Routes user intent to KB generation flow | Used to initiate KB draft creation and enforce minimal ambiguity handling |
| Scenario-to-KB | Converts a typed scenario into KB outputs | Applies output enforcement rules and style constraints |
| PDF Draft-to-KB | Converts a rough-draft PDF into KB outputs | Extracts + normalizes steps; fills gaps using grounded knowledge using AI Builder|
| Publish / Create KB Page | Sends SharePoint page content to backend | Designed for an Azure Function publish endpoint |
| Fallback / Ambiguity Handling | Asks minimum required question(s) | Only used when missing platform/version/scope blocks safe long-form output |

---

## 6. Outputs (Core Contract)

Article Forge’s key differentiator is its **strict three-output contract**.  This is 

### Output 1 — MSP KB ARTICLE (Markdown)

- Canonical KB article
- Adaptive Card-friendly Markdown
- Includes: Title, Summary, Symptoms, Root Cause (if known), Resolution (step-by-step), Verification, Notes for MSP, References

### Output 2 — SHAREPOINT PAGE CONTENT (Plain Text)

- End-user friendly, scannable copy
- Includes: Hero title, intro paragraph, key points, numbered steps, additional resources

### Output 3 — WORD DOCUMENT CONTENT (Template-ready Plain Text)

- Document Title, Purpose, Audience, High-Level Summary, Detailed Steps, Troubleshooting Notes, Revision History

These outputs are intended to be used as:

- **Review artifact** (Output 1)
- **Direct publishing payload** (Output 2)
- **Printable / formal distribution** (Output 3)

---

## 7. Actions (Azure Functions Integration)

Backend actions are used for publication and (optionally) structured storage.

### Representative Actions

#### 7.1 CreateKbDraft

- **Purpose**: Persist the generated KB draft (e.g., store metadata, assign ID)
- **Triggering rule**: Only invoked when the user explicitly confirms (e.g., selects Continue)

**Input schema (example):**

```json
{
  "pageTitle": "string",
  "siteUrl": "string",
  "htmlContent": "string",
  //"output1_markdown": "string",
  //"output2_sharepointText": "string",
  //"output3_wordText": "string",
}
```

#### 7.2 PublishKbToSharePoint

- **Purpose**: Create/update a SharePoint page using Output 2 (plain text)
- **Backend**: Azure Function → SharePoint page creation

-Call Microsoft Graph to generate the page, leave in draft form for human review before publishing

**Security model:**

- No secrets in Copilot Studio
- Backend authentication/authorization handled in Azure
- Copilot Studio to API: Easy Auth
- Logs avoid sensitive payloads; use correlation IDs

---

## 8. Security & Enterprise Readiness

- Deterministic output formats reduce drift and improve downstream automation reliability
- No tenant-specific identifiers or secrets included in generated content
- Publication is separated into a secured backend (Azure Functions)
- Supports governance through reviewable canonical Markdown and controlled publishing

---

## 9. Demo Evidence

Recommended repo artifacts:

- `/screenshots/` – example prompts, output samples, publish flow evidence
- `/architecture/` – one-page architecture diagram

(Optional) Demo video link: TBD

---
- **Back to ReadMe :** [`ReadMe.md`](../README.md)

