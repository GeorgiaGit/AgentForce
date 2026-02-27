![Article Forge](/assets/articleforge_small.png)

<p align="center">
  <img src="https://github.com/GeorgiaGit/AgentForce/assets/articleforge_small.png" alt="Description of Image"/>
</p>


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
- **Copilot Studio overview :** [`copilot-studio/agent-overview.md`](copilot-studio/agent-overview.md)
- **Architecture diagram:** [`architecture/architecture-diagram.mmd`](architecture/architecture-diagram.mmd)  
- **Azure Function Overview:** [`VSCode/copilot-easy-auth-alignment-runbook.md`](VSCode/copilot-easy-auth-alignment-runbook.md)  
  (Optional PNG export: `architecture/architecture-diagram.png`)

### Security note (public repo)
This repository intentionally excludes Copilot Studio solution exports and any tenant-specific identifiers or secrets. Configuration and authentication are handled via Azure configuration (no secrets committed).
