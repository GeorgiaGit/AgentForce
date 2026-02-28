# Public Azure Function Template (Sanitized)

This folder contains a minimal, sanitized Azure Functions project for sharing in a public repository.

## Included files
- `CreateKBDrafts_Public.cs` (sanitized function implementation)
- `Program.cs`
- `host.json`
- `SharePointPageCreation.Public.csproj`
- `payload.public.json`
- `test-call.public.ps1`

## What was sanitized
- Tenant/site-specific URLs and paths
- Function endpoint URL and function key
- Internal naming conventions and environment-specific identifiers
- Verbose logs that could expose sensitive payload details

## Placeholder key
Use these values consistently across Azure CLI commands, connector config, and test calls.

**Azure infrastructure**
- `<SUBSCRIPTION_ID>`: target Azure subscription
- `<RESOURCE_GROUP>`: resource group hosting the function app
- `<AZURE_REGION>`: deployment region (for example `eastus`)
- `<STORAGE_ACCOUNT_NAME>`: storage account used by the function app
- `<FUNCTION_APP_NAME>`: Azure Function App resource name
- `<FUNCTION_HOSTNAME>`: public host (for example `<FUNCTION_APP_NAME>.azurewebsites.net`)

**Auth and API identity**
- `<FUNCTION_KEY>`: function access key for HTTP trigger calls
- `<API_APP_CLIENT_ID>`: API app registration client ID used by Easy Auth
- `<COPILOT_CLIENT_APP_ID>`: Copilot client app registration ID (allowed caller)
- `<API_IDENTIFIER_URI>`: API audience/identifier URI (for example `api://<API_APP_CLIENT_ID>`)

For full identity mapping and Easy Auth alignment, see [copilot-easy-auth-alignment-runbook.md](copilot-easy-auth-alignment-runbook.md).

## Prerequisites
- .NET 8 SDK
- Azure Functions Core Tools v4
- Azure CLI authenticated to your subscription

## Local run
1. Create your own `local.settings.json` (do not commit):
   ```json
   {
     "IsEncrypted": false,
     "Values": {
       "AzureWebJobsStorage": "UseDevelopmentStorage=true",
       "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
       "DEBUG_RETURN_ERRORS": "false"
     }
   }
   ```
2. Build:
   ```powershell
   dotnet build .\SharePointPageCreation.Public.csproj
   ```
3. Start:
   ```powershell
   func start
   ```

## Deploy
```powershell
dotnet publish .\SharePointPageCreation.Public.csproj -c Release
func azure functionapp publish <FUNCTION_APP_NAME>
```

## Deployment quickstart (Azure)
Use this when you want to stand up and publish quickly, then move to full auth hardening.

1. Login and select subscription:
   ```powershell
   az login
   az account set --subscription <SUBSCRIPTION_ID>
   ```

2. Create core resources:
   ```powershell
   az group create --name <RESOURCE_GROUP> --location <AZURE_REGION>
   az storage account create --name <STORAGE_ACCOUNT_NAME> --resource-group <RESOURCE_GROUP> --location <AZURE_REGION> --sku Standard_LRS
   az functionapp create --name <FUNCTION_APP_NAME> --resource-group <RESOURCE_GROUP> --consumption-plan-location <AZURE_REGION> --storage-account <STORAGE_ACCOUNT_NAME> --runtime dotnet-isolated --functions-version 4
   ```

3. Configure app settings and managed identity:
   ```powershell
   az functionapp config appsettings set --name <FUNCTION_APP_NAME> --resource-group <RESOURCE_GROUP> --settings FUNCTIONS_WORKER_RUNTIME=dotnet-isolated DEBUG_RETURN_ERRORS=false
   az functionapp identity assign --name <FUNCTION_APP_NAME> --resource-group <RESOURCE_GROUP>
   ```

4. Publish code:
   ```powershell
   dotnet publish .\SharePointPageCreation.Public.csproj -c Release
   func azure functionapp publish <FUNCTION_APP_NAME>
   ```

5. Run smoke test:
   ```powershell
   .\test-call.public.ps1 -FunctionUrl "https://<FUNCTION_HOSTNAME>/api/CreateKbDraft_Public?code=<FUNCTION_KEY>"
   ```

For complete setup (Easy Auth alignment, Entra app registrations, allowed audiences/apps, and troubleshooting), see:
- [copilot-easy-auth-alignment-runbook.md](copilot-easy-auth-alignment-runbook.md)

## Test call
```powershell
.\test-call.public.ps1 -FunctionUrl "https://<FUNCTION_HOSTNAME>/api/CreateKbDraft_Public?code=<FUNCTION_KEY>"
```

## Pull request instructions
Use this workflow when opening a PR to a public repository.

1. Sync and create a feature branch:
   ```powershell
   git checkout main
   git pull
   git checkout -b chore/add-sharepoint-pagecreation-template
   ```

2. Validate before commit:
   - Confirm no secrets are present (`local.settings.json` must not be committed).
   - Confirm build output folders are excluded (`bin/`, `obj/`).
   - Build the project:
     ```powershell
     dotnet build .\SharePointPageCreation.Public.csproj
     ```

3. Commit and push:
   ```powershell
   git add .
   git commit -m "Add sanitized SharePoint page creation Azure Function template"
   git push -u origin chore/add-sharepoint-pagecreation-template
   ```

4. Open the PR with this checklist:
   - [ ] Project builds successfully.
   - [ ] No secrets, tenant identifiers, or real function keys included.
   - [ ] Sample payload and URLs use fake/public-safe values.
   - [ ] README updated for setup, deploy, and test usage.

Suggested PR title:
- `Add sanitized SharePoint page creation Azure Function template`

Suggested PR description:
- Adds a public-safe Azure Functions template for SharePoint page draft creation.
- Replaces internal identifiers with fake valid values.
- Includes sample payload, deployment/test commands, and public repo guidance.
