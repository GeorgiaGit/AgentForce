
## Copilot Studio Agent (Article Forge)

**Article Forge** is a deterministic Knowledge Base (KB) documentation generator built in **Microsoft Copilot Studio**, with a secure **Azure Functions** publishing boundary.

It transforms either:
- a typed scenario (e.g., “Clear Teams cache on Windows and macOS”), or
- an uploaded rough-draft PDF

into **publication-ready KB content** using a strict three-output contract:
1) MSP KB Article (Markdown — canonical review version)  
2) SharePoint Page Content (plain text — publishing payload)  
3) Word Document Content (template-ready plain text)

### Key docs
- **Copilot Studio overview (judge-ready):** [`copilot-studio/agent-overview.md`](copilot-studio/agent-overview.md)
- **Architecture diagram:** [`architecture/architecture-diagram.mmd`](architecture/architecture-diagram.mmd)  
  (Optional PNG export: `architecture/architecture-diagram.png`)

### Security note (public repo)
This repository intentionally excludes Copilot Studio solution exports and any tenant-specific identifiers or secrets. Configuration and authentication are handled via Azure configuration (no secrets committed).
