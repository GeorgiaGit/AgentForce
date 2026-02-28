
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace Company.Function;

public class CreateKbDraftPublic
{
    private static readonly HttpClient Http = new HttpClient();
    private const string DefaultTemplateServerRelativePath = "/sites/demo-public/SitePages/Templates/template-page.aspx";

    [Function("CreateKbDraft_Public")]
    public static async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req,
        FunctionContext ctx)
    {
        var log = ctx.GetLogger("CreateKbDraft");
        var correlationId = ctx.InvocationId;

        var debugReturnErrors =
            string.Equals(Environment.GetEnvironmentVariable("DEBUG_RETURN_ERRORS"), "true", StringComparison.OrdinalIgnoreCase);
        log.LogInformation("DEBUG_RETURN_ERRORS={Debug}", debugReturnErrors);

        try
        {
            // ---- Read & validate request ----
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var input = JsonSerializer.Deserialize<CreateKbDraftRequest>(requestBody, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (input is null ||
                string.IsNullOrWhiteSpace(input.SiteUrl) ||
                string.IsNullOrWhiteSpace(input.PageTitle) ||
                string.IsNullOrWhiteSpace(input.HtmlContent))
            {
                return await Error(req, HttpStatusCode.BadRequest, "BadRequest",
                    "Missing required fields: siteUrl, pageTitle, htmlContent.", correlationId);
            }

            var siteUrl = input.SiteUrl;
            var pageTitle = input.PageTitle;
            var htmlContent = input.HtmlContent;

            log.LogInformation(
                "CreateKbDraft request received. SiteHost={SiteHost}, PageTitleLength={PageTitleLength}, FileNameProvided={FileNameProvided}",
                SafeHostFromUrl(siteUrl),
                pageTitle.Length,
                !string.IsNullOrWhiteSpace(input.FileName));

            // ---- Acquire Graph token (Managed Identity in Azure) ----
            var credential = new DefaultAzureCredential();
            var token = await credential.GetTokenAsync(
                new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" })
            );

            Http.DefaultRequestHeaders.Clear();
            Http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);

            // Recommended header when reusing responses in subsequent requests (reduces metadata noise)
            Http.DefaultRequestHeaders.Accept.Clear();
            Http.DefaultRequestHeaders.Accept.Add(
                new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            Http.DefaultRequestHeaders.Add("Accept", "application/json;odata.metadata=none");

            // ---- Resolve siteId ----
            var siteId = await ResolveSiteIdAsync(siteUrl, log);

            var createdPageId = await CreatePageWithUniqueNameAsync(siteId, pageTitle, input.FileName, log);

            await EnsureOneColumnCanvasLayoutAsync(siteId, createdPageId, log);

            var pageWithLayoutJson = await GetPageWithCanvasLayoutAsync(siteId, createdPageId, log);
            var (sectionId, columnId) = ExtractFirstSectionAndColumnIds(pageWithLayoutJson);

            if (string.IsNullOrWhiteSpace(sectionId) || string.IsNullOrWhiteSpace(columnId))
            {
                log.LogError("Unable to find sectionId/columnId after PATCH.");
                return await Error(req, HttpStatusCode.InternalServerError, "InternalServerError",
                    "Unable to find horizontal section/column IDs after setting canvasLayout.", correlationId);
            }

            await AddTextWebPartAsync(siteId, createdPageId, sectionId!, columnId!, htmlContent, log);

            var finalJson = await GetPageWithCanvasLayoutAsync(siteId, createdPageId, log);
            var (pageId, draftPageUrl) = TryExtractPageIdAndWebUrl(finalJson);

            var ok = req.CreateResponse(HttpStatusCode.OK);
            ok.Headers.Add("Content-Type", "application/json");

            if (!string.IsNullOrWhiteSpace(pageId))
                ok.Headers.Add("X-Draft-Page-Id", pageId);

            if (!string.IsNullOrWhiteSpace(draftPageUrl))
            {
                ok.Headers.Add("X-Draft-Page-Url", draftPageUrl);
                log.LogInformation("Draft page created. PageId={PageId}, WebUrl={WebUrl}", pageId, draftPageUrl);
            }

            var successPayload = new
            {
                success = true,
                pageUrl = draftPageUrl ?? string.Empty,
                serverRelativeUrl = TryGetServerRelativeUrl(draftPageUrl),
                message = "Draft page created successfully."
            };

            await ok.WriteStringAsync(JsonSerializer.Serialize(successPayload));

            return ok;
        }
        catch (GraphApiException graphEx)
        {
            log.LogError(graphEx,
                "Graph call failed. Status={StatusCode}, GraphCode={GraphCode}, RetryAfterSeconds={RetryAfterSeconds}, RequestId={RequestId}, Method={Method}, Url={Url}",
                (int)graphEx.StatusCode,
                graphEx.GraphCode,
                graphEx.RetryAfterSeconds,
                graphEx.RequestId,
                graphEx.Method,
                graphEx.Url);

            var message = debugReturnErrors
                ? $"{graphEx.GraphMessage ?? "Graph request failed."} RequestId={graphEx.RequestId ?? "n/a"}."
                : (graphEx.GraphMessage ?? "Graph request failed.");

            return await Error(req, graphEx.StatusCode, graphEx.GraphCode ?? "GraphError", message, correlationId);
        }
        catch (Exception ex)
        {
            log.LogError(ex, "Unhandled exception in CreateKbDraft.");

            var message = debugReturnErrors ? ex.Message : "Internal Server Error. Check function logs.";
            return await Error(req, HttpStatusCode.InternalServerError, "InternalServerError", message, correlationId);
        }
    }

    // -------------------------
    // Graph helper operations
    // -------------------------

    private static async Task<string> ResolveSiteIdAsync(string siteUrl, ILogger log)
    {
        var siteUri = new Uri(siteUrl);
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:{siteUri.AbsolutePath}";
        var body = await SendGraphAsync(HttpMethod.Get, url, null, log);
        return JsonDocument.Parse(body).RootElement.GetProperty("id").GetString()
               ?? throw new InvalidOperationException("Unable to resolve siteId from siteUrl.");
    }

    /// <summary>
    /// IMPORTANT: Must cast to microsoft.graph.sitePage to expand canvasLayout. [3](https://www.linkedin.com/pulse/how-extend-copilot-studio-external-apis-real-world-automation-wali-ir4nf)
    /// </summary>
    private static async Task<string> GetPageWithCanvasLayoutAsync(string siteId, string pageId, ILogger log)
    {
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage?$expand=canvasLayout";
        return await SendGraphAsync(HttpMethod.Get, url, null, log);
    }

    /// <summary>
    /// IMPORTANT: Update must be PATCH .../microsoft.graph.sitePage, not PATCH .../pages/{pageId}. [1](https://www.linkedin.com/pulse/how-use-azure-function-mcp-servers-copilot-studio-dtower-software-uj1wf)[2](https://github.com/microsoft/CopilotStudioSamples/tree/main/CallAgentConnector)
    /// </summary>
    private static async Task EnsureOneColumnCanvasLayoutAsync(string siteId, string pageId, ILogger log)
    {
        var patchPayload = new Dictionary<string, object?>
        {
            ["@odata.type"] = "#microsoft.graph.sitePage",
            ["canvasLayout"] = new
            {
                horizontalSections = new[]
                {
                    new
                    {
                        layout = "oneColumn",
                        columns = new[]
                        {
                            new { width = 12 }
                        }
                    }
                }
            }
        };

        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage";
        _ = await SendGraphAsync(new HttpMethod("PATCH"), url, patchPayload, log);
    }

    private static async Task AddTextWebPartAsync(
        string siteId,
        string pageId,
        string sectionId,
        string columnId,
        string html,
        ILogger log)
    {
        var payload = new Dictionary<string, object?>
        {
            ["@odata.type"] = "#microsoft.graph.textwebpart",
            ["innerHtml"] = html
        };

        // Horizontal section/column webparts endpoint documented for Create webPart. [4](https://www.youtube.com/watch?v=uo-vCFL96yQ)
        var url =
            $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage/canvasLayout/horizontalSections/{sectionId}/columns/{columnId}/webparts";

        _ = await SendGraphAsync(HttpMethod.Post, url, payload, log);
    }

    /// <summary>
    /// Central Graph sender with debug-first logging and consistent error bubbling.
    /// Always returns response body if success; throws InvalidOperationException with body if non-success.
    /// </summary>
    private static async Task<string> SendGraphAsync(HttpMethod method, string url, object? payload, ILogger log)
    {
        const int maxAttempts = 3;

        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            log.LogInformation("{Method} {Url} (attempt {Attempt}/{MaxAttempts})", method.Method, url, attempt, maxAttempts);

            using var req = new HttpRequestMessage(method, url);

            if (payload is not null)
            {
                var json = JsonSerializer.Serialize(payload);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
            }

            using var resp = await Http.SendAsync(req);
            var body = await resp.Content.ReadAsStringAsync();

            log.LogInformation("Graph status={StatusCode} for {Method} {Url}", (int)resp.StatusCode, method.Method, url);

            if (resp.IsSuccessStatusCode)
                return body;

            var isTransient = IsTransientGraphStatus(resp.StatusCode);
            if (isTransient && attempt < maxAttempts)
            {
                var delay = GetRetryDelay(resp, body, attempt);
                log.LogWarning("Transient Graph error status={StatusCode}. Retrying after {DelaySeconds}s.",
                    (int)resp.StatusCode,
                    Math.Round(delay.TotalSeconds, 2));

                await Task.Delay(delay);
                continue;
            }

            log.LogError("Graph error body length: {BodyLength}", body?.Length ?? 0);
            throw BuildGraphApiException(method, url, resp, body, attempt, maxAttempts);
        }

        throw new InvalidOperationException("Graph call failed after retries.");
    }

    private static GraphApiException BuildGraphApiException(
        HttpMethod method,
        string url,
        HttpResponseMessage response,
        string? responseBody,
        int attempt,
        int maxAttempts)
    {
        var safeResponseBody = responseBody ?? string.Empty;
        var graphCode = TryGetGraphErrorField(safeResponseBody, "code");
        var graphMessage = TryGetGraphErrorField(safeResponseBody, "message");

        var retryAfterSeconds = ReadRetryAfterSeconds(response, safeResponseBody);
        var requestId = ReadHeaderValue(response, "request-id") ?? TryGetGraphInnerErrorField(safeResponseBody, "request-id");
        var clientRequestId = ReadHeaderValue(response, "client-request-id") ?? TryGetGraphInnerErrorField(safeResponseBody, "client-request-id");

        var message = $"Graph API call failed. Status={(int)response.StatusCode}, Code={graphCode ?? "unknown"}, RequestId={requestId ?? "n/a"}, Url={url}";

        return new GraphApiException(
            message,
            response.StatusCode,
            method.Method,
            url,
            graphCode,
            graphMessage,
            safeResponseBody,
            requestId,
            clientRequestId,
            retryAfterSeconds,
            attempt,
            maxAttempts);
    }

    private static bool IsTransientGraphStatus(HttpStatusCode statusCode)
    {
        return statusCode == HttpStatusCode.TooManyRequests
            || statusCode == HttpStatusCode.ServiceUnavailable
            || statusCode == HttpStatusCode.GatewayTimeout;
    }

    private static TimeSpan GetRetryDelay(HttpResponseMessage response, string body, int attempt)
    {
        var seconds = ReadRetryAfterSeconds(response, body);

        if (seconds == 0)
            seconds = Math.Min(12, (int)Math.Pow(2, attempt + 1));

        seconds = Math.Min(seconds, 30);
        var jitterMs = Random.Shared.Next(50, 400);
        return TimeSpan.FromSeconds(seconds) + TimeSpan.FromMilliseconds(jitterMs);
    }

    private static int ReadRetryAfterSeconds(HttpResponseMessage response, string body)
    {
        if (response.Headers.TryGetValues("Retry-After", out var retryAfterValues))
        {
            var retryAfterValue = retryAfterValues.FirstOrDefault();
            if (int.TryParse(retryAfterValue, out var headerSeconds) && headerSeconds > 0)
                return headerSeconds;
        }

        var bodyRetryAfter = TryGetGraphErrorField(body, "retryAfterSeconds");
        if (int.TryParse(bodyRetryAfter, out var bodySeconds) && bodySeconds > 0)
            return bodySeconds;

        return 0;
    }

    private static string? TryGetGraphErrorField(string body, string fieldName)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("error", out var error) &&
                error.TryGetProperty(fieldName, out var value))
            {
                return value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
            }
        }
        catch
        {
            // Ignore parse errors.
        }

        return null;
    }

    private static string? TryGetGraphInnerErrorField(string body, string fieldName)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("error", out var error) &&
                error.TryGetProperty("innerError", out var innerError) &&
                innerError.TryGetProperty(fieldName, out var value))
            {
                return value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
            }
        }
        catch
        {
            // Ignore parse errors.
        }

        return null;
    }

    private static string? ReadHeaderValue(HttpResponseMessage response, string headerName)
    {
        if (response.Headers.TryGetValues(headerName, out var values))
            return values.FirstOrDefault();

        return null;
    }

    // -------------------------
    // Name helpers (dedupe-safe)
    // -------------------------

    private static async Task<string> CreateDraftFromTemplateAsync(
        string siteId,
        string siteRelativePrefix,
        string templateServerRelativePath,
        string pageTitle,
        string? fileNameOverride,
        string htmlContent,
        ILogger log)
    {
        var templatePageId = await ResolveTemplatePageIdAsync(siteId, siteRelativePrefix, templateServerRelativePath, log);
        var templatePageJson = await GetTemplatePageWithCanvasLayoutAsync(siteId, templatePageId, log);

        var templateCanvasLayout = ExtractCanvasLayoutSanitized(templatePageJson);
        var injectedCanvasLayout = InjectHtmlIntoFirstTextWebPart(templateCanvasLayout, htmlContent, log);

        var createdPageId = await CreatePageFromCanvasLayoutWithUniqueNameAsync(
            siteId,
            pageTitle,
            fileNameOverride,
            injectedCanvasLayout,
            log);

        return await GetPageWithCanvasLayoutAsync(siteId, createdPageId, log);
    }

    private static async Task<string> ResolveTemplatePageIdAsync(
        string siteId,
        string siteRelativePrefix,
        string templateServerRelativePath,
        ILogger log)
    {
        var templateName = ExtractTemplateFileName(templateServerRelativePath);
        if (!string.IsNullOrWhiteSpace(templateName))
        {
            try
            {
                var pageId = await GetPageIdByNameAsync(siteId, templateName, log);
                log.LogInformation("Resolved template page by name. TemplateName={TemplateName}, PageId={PageId}", templateName, pageId);
                return pageId;
            }
            catch (Exception ex)
            {
                log.LogWarning(ex,
                    "Template lookup by pages collection failed. Falling back to Site Pages list lookup. TemplateName={TemplateName}",
                    templateName);
            }
        }

        try
        {
            var fallbackId = await GetTemplatePageIdFromSitePagesListAsync(siteId, siteRelativePrefix, templateServerRelativePath, log);
            log.LogInformation("Resolved template page via Site Pages list fallback. PageId={PageId}", fallbackId);
            return fallbackId;
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Template lookup by Site Pages list failed.");
        }

        throw new InvalidOperationException(
            "Unable to resolve template page id from Graph pages collection or Site Pages list. " +
            "Set templateServerRelativePath to a valid page under SitePages/Templates and verify app permissions.");
    }

    private static string ExtractTemplateFileName(string templateServerRelativePath)
    {
        var normalized = NormalizeTemplatePath(templateServerRelativePath);
        if (string.IsNullOrWhiteSpace(normalized))
            return string.Empty;

        var unescaped = Uri.UnescapeDataString(normalized.Trim());
        var parts = unescaped.Split('/', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? string.Empty : parts[^1];
    }

    private static async Task<string> GetPageIdByNameAsync(string siteId, string pageName, ILogger log)
    {
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages?$select=id,name&$top=200";

        while (!string.IsNullOrWhiteSpace(url))
        {
            var body = await SendGraphAsync(HttpMethod.Get, url, null, log);

            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("value", out var value) || value.ValueKind != JsonValueKind.Array)
                throw new InvalidOperationException("Graph pages response did not contain a value array.");

            foreach (var item in value.EnumerateArray())
            {
                var name = item.TryGetProperty("name", out var n) ? n.GetString() : null;
                if (!string.Equals(name, pageName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var id = item.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
                if (!string.IsNullOrWhiteSpace(id))
                    return id!;
            }

            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var nextLink)
                ? nextLink.GetString() ?? string.Empty
                : string.Empty;
        }

        throw new InvalidOperationException($"Template page '{pageName}' was not found in Site Pages.");
    }

    private static async Task<string> GetListItemUniqueIdFromSitePagesPathAsync(
        string siteId,
        string siteRelativePrefix,
        string serverRelativePath,
        ILogger log)
    {
        var driveRelativePath = ToDriveRelativePath(siteRelativePrefix, serverRelativePath);
        var encodedPath = EncodeServerRelativePathForGraph(driveRelativePath);
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:{encodedPath}";

        var body = await SendGraphAsync(HttpMethod.Get, url, null, log);
        var root = JsonDocument.Parse(body).RootElement;

        var uniqueId = root.GetProperty("sharepointIds").GetProperty("listItemUniqueId").GetString();
        if (string.IsNullOrWhiteSpace(uniqueId))
            throw new InvalidOperationException("Template listItemUniqueId not found in driveItem response.");

        return uniqueId!;
    }

    private static async Task<string> GetTemplatePageIdFromSitePagesListAsync(
        string siteId,
        string siteRelativePrefix,
        string templateServerRelativePath,
        ILogger log)
    {
        var templatePath = NormalizeTemplatePath(templateServerRelativePath);
        var templatePathUnescaped = Uri.UnescapeDataString(templatePath).Trim();
        var templateName = ExtractTemplateFileName(templatePath);

        var sitePagesListId = await GetSitePagesListIdAsync(siteId, log);

        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{sitePagesListId}/items?$expand=fields($select=FileLeafRef,FileRef,UniqueId)&$top=200";

        while (!string.IsNullOrWhiteSpace(url))
        {
            var body = await SendGraphAsync(HttpMethod.Get, url, null, log);
            using var doc = JsonDocument.Parse(body);

            if (doc.RootElement.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in value.EnumerateArray())
                {
                    if (!item.TryGetProperty("fields", out var fields) || fields.ValueKind != JsonValueKind.Object)
                        continue;

                    var fileLeafRef = fields.TryGetProperty("FileLeafRef", out var leaf) ? leaf.GetString() : null;
                    var fileRef = fields.TryGetProperty("FileRef", out var fileRefEl) ? fileRefEl.GetString() : null;

                    var nameMatches = !string.IsNullOrWhiteSpace(fileLeafRef) &&
                        !string.IsNullOrWhiteSpace(templateName) &&
                        string.Equals(fileLeafRef, templateName, StringComparison.OrdinalIgnoreCase);

                    var pathMatches = PathMatchesTemplate(siteRelativePrefix, fileRef, templatePathUnescaped);

                    if (!nameMatches && !pathMatches)
                        continue;

                    var uniqueId = fields.TryGetProperty("UniqueId", out var uniqueIdEl)
                        ? uniqueIdEl.GetString()
                        : null;

                    if (!string.IsNullOrWhiteSpace(uniqueId))
                        return uniqueId!;

                    var listItemId = item.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(listItemId))
                        return listItemId!;
                }
            }

            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var nextLink)
                ? nextLink.GetString() ?? string.Empty
                : string.Empty;
        }

        throw new InvalidOperationException(
            $"Template page '{templatePathUnescaped}' was not found in Site Pages list.");
    }

    private static async Task<string> GetSitePagesListIdAsync(string siteId, ILogger log)
    {
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists?$select=id,displayName,webUrl,list,system";
        var body = await SendGraphAsync(HttpMethod.Get, url, null, log);

        using var doc = JsonDocument.Parse(body);
        if (!doc.RootElement.TryGetProperty("value", out var value) || value.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("Unable to list site lists.");

        var seen = new List<string>();
        foreach (var list in value.EnumerateArray())
        {
            var displayName = list.TryGetProperty("displayName", out var nameEl) ? nameEl.GetString() : null;
            var webUrl = list.TryGetProperty("webUrl", out var webUrlEl) ? webUrlEl.GetString() : null;
            var template = list.TryGetProperty("list", out var listObj) &&
                           listObj.ValueKind == JsonValueKind.Object &&
                           listObj.TryGetProperty("template", out var templateEl)
                ? templateEl.GetString()
                : null;
            var isSystemList = list.TryGetProperty("system", out var systemObj) && systemObj.ValueKind == JsonValueKind.Object;

            seen.Add($"{displayName ?? "(null)"}|{template ?? "(null)"}|{webUrl ?? "(null)"}|system={isSystemList}");

            if (!IsLikelySitePagesList(displayName, webUrl, template))
                continue;

            var id = list.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            if (!string.IsNullOrWhiteSpace(id))
                return id!;
        }

        throw new InvalidOperationException($"Site Pages list not found. Lists seen: {string.Join("; ", seen)}");
    }

    private static bool IsLikelySitePagesList(string? displayName, string? webUrl, string? template)
    {
        if (!string.IsNullOrWhiteSpace(template))
        {
            if (template.Equals("sitePageLibrary", StringComparison.OrdinalIgnoreCase) ||
                template.Equals("webPageLibrary", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (!string.IsNullOrWhiteSpace(displayName) &&
            (displayName.Equals("Site Pages", StringComparison.OrdinalIgnoreCase) ||
             displayName.Equals("Pages", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(webUrl) && webUrl.Contains("/SitePages", StringComparison.OrdinalIgnoreCase))
            return true;

        return false;
    }

    private static bool PathMatchesTemplate(string siteRelativePrefix, string? fileRef, string templatePathUnescaped)
    {
        if (string.IsNullOrWhiteSpace(fileRef))
            return false;

        var normalizedFileRef = Uri.UnescapeDataString(fileRef).Trim();
        var normalizedTemplate = Uri.UnescapeDataString(templatePathUnescaped ?? string.Empty).Trim();

        if (!normalizedTemplate.StartsWith('/'))
            normalizedTemplate = "/" + normalizedTemplate;

        if (string.Equals(normalizedFileRef, normalizedTemplate, StringComparison.OrdinalIgnoreCase))
            return true;

        var driveRelativeTemplate = ToDriveRelativePath(siteRelativePrefix, normalizedTemplate);
        return string.Equals(normalizedFileRef, driveRelativeTemplate, StringComparison.OrdinalIgnoreCase);
    }

    private static string EncodeServerRelativePathForGraph(string serverRelativePath)
    {
        var normalized = Uri.UnescapeDataString(serverRelativePath ?? string.Empty).Trim();
        if (!normalized.StartsWith('/'))
            normalized = "/" + normalized;

        var encodedSegments = normalized
            .Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(Uri.EscapeDataString);

        return "/" + string.Join("/", encodedSegments);
    }

    private static string NormalizeTemplatePath(string templatePathInput)
    {
        var value = (templatePathInput ?? string.Empty).Trim();
        if (Uri.TryCreate(value, UriKind.Absolute, out var absolute))
            return absolute.AbsolutePath;

        return value;
    }

    private static string ToDriveRelativePath(string siteRelativePrefix, string serverRelativePath)
    {
        var normalizedPrefix = Uri.UnescapeDataString(siteRelativePrefix ?? string.Empty).Trim().TrimEnd('/');
        var normalizedPath = Uri.UnescapeDataString(serverRelativePath ?? string.Empty).Trim();

        if (!normalizedPath.StartsWith('/'))
            normalizedPath = "/" + normalizedPath;

        if (!string.IsNullOrWhiteSpace(normalizedPrefix) &&
            normalizedPath.StartsWith(normalizedPrefix + "/", StringComparison.OrdinalIgnoreCase))
        {
            return normalizedPath.Substring(normalizedPrefix.Length);
        }

        return normalizedPath;
    }

    private static async Task<string> GetTemplatePageWithCanvasLayoutAsync(string siteId, string listItemUniqueId, ILogger log)
    {
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages/{listItemUniqueId}/microsoft.graph.sitePage?$expand=canvasLayout";
        return await SendGraphAsync(HttpMethod.Get, url, null, log);
    }

    private static JsonElement ExtractCanvasLayoutSanitized(string templatePageJson)
    {
        using var doc = JsonDocument.Parse(templatePageJson);
        var root = doc.RootElement;

        if (!root.TryGetProperty("canvasLayout", out var canvasLayout))
            throw new InvalidOperationException("Template page does not contain canvasLayout.");

        var sanitizedJson = SanitizeJsonElement(canvasLayout);

        using var sanitizedDoc = JsonDocument.Parse(sanitizedJson);
        return sanitizedDoc.RootElement.Clone();
    }

    private static string SanitizeJsonElement(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            var dict = new Dictionary<string, object?>();
            foreach (var prop in element.EnumerateObject())
            {
                var name = prop.Name;

                if (name.EndsWith("@odata.context", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (name.StartsWith("@odata.", StringComparison.OrdinalIgnoreCase) &&
                    !name.Equals("@odata.type", StringComparison.OrdinalIgnoreCase))
                    continue;

                dict[name] = JsonElementToObject(prop.Value);
            }
            return JsonSerializer.Serialize(dict);
        }

        if (element.ValueKind == JsonValueKind.Array)
        {
            var list = element.EnumerateArray().Select(JsonElementToObject).ToList();
            return JsonSerializer.Serialize(list);
        }

        return element.GetRawText();
    }

    private static object? JsonElementToObject(JsonElement el)
    {
        return el.ValueKind switch
        {
            JsonValueKind.Object => JsonSerializer.Deserialize<Dictionary<string, object?>>(SanitizeJsonElement(el)),
            JsonValueKind.Array => el.EnumerateArray().Select(JsonElementToObject).ToList(),
            JsonValueKind.String => el.GetString(),
            JsonValueKind.Number => el.TryGetInt64(out var l) ? l : el.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => null
        };
    }

    private static JsonElement InjectHtmlIntoFirstTextWebPart(JsonElement canvasLayout, string html, ILogger log)
    {
        bool replaced = false;
        var mutated = MutateFirstTextWebPart(canvasLayout, html, ref replaced);

        if (!replaced)
            log.LogWarning("No textWebPart found in template canvasLayout. No injection performed.");

        var json = JsonSerializer.Serialize(mutated);
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static object? MutateFirstTextWebPart(JsonElement element, string html, ref bool replaced)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            var dict = new Dictionary<string, object?>();

            foreach (var prop in element.EnumerateObject())
            {
                if (prop.NameEquals("webparts") && prop.Value.ValueKind == JsonValueKind.Array)
                {
                    var newWebparts = new List<object?>();

                    foreach (var wp in prop.Value.EnumerateArray())
                    {
                        var isTextWebPart =
                            wp.ValueKind == JsonValueKind.Object &&
                            wp.TryGetProperty("@odata.type", out var t) &&
                            (t.GetString()?.Equals("#microsoft.graph.textWebPart", StringComparison.OrdinalIgnoreCase) == true ||
                             t.GetString()?.Equals("#microsoft.graph.textwebpart", StringComparison.OrdinalIgnoreCase) == true);

                        // Keep only Graph-supported text webparts when creating from template canvas.
                        // Template pages often contain first-party webparts not accepted by Graph create.
                        if (!isTextWebPart)
                            continue;

                        if (!replaced)
                        {
                            var wpDict = wp.EnumerateObject().ToDictionary(p => p.Name, p => JsonElementToObject(p.Value));
                            wpDict["innerHtml"] = html;
                            newWebparts.Add(wpDict);
                            replaced = true;
                        }
                        else
                        {
                            newWebparts.Add(JsonElementToObject(wp));
                        }
                    }

                    dict[prop.Name] = newWebparts;
                }
                else
                {
                    dict[prop.Name] = MutateFirstTextWebPart(prop.Value, html, ref replaced);
                }
            }

            return dict;
        }

        if (element.ValueKind == JsonValueKind.Array)
        {
            var list = new List<object?>();
            foreach (var e in element.EnumerateArray())
                list.Add(MutateFirstTextWebPart(e, html, ref replaced));
            return list;
        }

        return JsonElementToObject(element);
    }

    private static async Task<string> CreatePageFromCanvasLayoutWithUniqueNameAsync(
        string siteId,
        string pageTitle,
        string? fileNameOverride,
        JsonElement canvasLayout,
        ILogger log)
    {
        var baseName = BuildBaseName(pageTitle, fileNameOverride);

        for (int i = 0; i < 20; i++)
        {
            var name = i == 0 ? $"{baseName}.aspx" : $"{baseName}-{i}.aspx";

            var payload = new Dictionary<string, object?>
            {
                ["@odata.type"] = "#microsoft.graph.sitePage",
                ["name"] = name,
                ["title"] = pageTitle,
                ["pageLayout"] = "article",
                ["canvasLayout"] = JsonSerializer.Deserialize<object>(canvasLayout.GetRawText())
            };

            var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages";

            try
            {
                var body = await SendGraphAsync(HttpMethod.Post, url, payload, log);
                var id = JsonDocument.Parse(body).RootElement.GetProperty("id").GetString();

                if (string.IsNullOrWhiteSpace(id))
                    throw new InvalidOperationException("Create page succeeded but no id returned.");

                log.LogInformation("Created template-based page successfully. name={Name}, id={Id}", name, id);
                return id!;
            }
            catch (InvalidOperationException ex)
            {
                if (ex.Message.Contains("\"code\":\"nameAlreadyExists\"", StringComparison.OrdinalIgnoreCase))
                {
                    log.LogWarning("Name already exists (template create), retrying with suffix. name={Name}", name);
                    continue;
                }
                throw;
            }
        }

        throw new InvalidOperationException("Unable to create a unique template-based page name after multiple attempts.");
    }

    private static async Task<string> CreatePageWithUniqueNameAsync(
        string siteId,
        string pageTitle,
        string? fileNameOverride,
        ILogger log)
    {
        var baseName = BuildBaseName(pageTitle, fileNameOverride);

        for (int i = 0; i < 20; i++)
        {
            var name = i == 0 ? $"{baseName}.aspx" : $"{baseName}-{i}.aspx";

            var payload = new Dictionary<string, object?>
            {
                ["@odata.type"] = "#microsoft.graph.sitePage",
                ["name"] = name,
                ["title"] = pageTitle,
                ["pageLayout"] = "article"
            };

            var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages";

            try
            {
                var body = await SendGraphAsync(HttpMethod.Post, url, payload, log);
                var id = JsonDocument.Parse(body).RootElement.GetProperty("id").GetString();

                if (string.IsNullOrWhiteSpace(id))
                    throw new InvalidOperationException("Create page succeeded but no id returned.");

                log.LogInformation("Created page successfully. name={Name}, id={Id}", name, id);
                return id!;
            }
            catch (InvalidOperationException ex)
            {
                if (ex.Message.Contains("\"code\":\"nameAlreadyExists\"", StringComparison.OrdinalIgnoreCase))
                {
                    log.LogWarning("Name already exists, retrying with suffix. name={Name}", name);
                    continue;
                }
                throw;
            }
        }

        throw new InvalidOperationException("Unable to create a unique page name after multiple attempts.");
    }

    private static string BuildBaseName(string pageTitle, string? fileNameOverride)
    {
        if (!string.IsNullOrWhiteSpace(fileNameOverride))
        {
            var stem = Regex.Replace(fileNameOverride.Trim(), @"\.aspx$", "", RegexOptions.IgnoreCase);
            stem = Regex.Replace(stem, @"\s+", "-");
            stem = Regex.Replace(stem, @"[^A-Za-z0-9\-]", "");
            stem = Regex.Replace(stem, @"\-{2,}", "-").Trim('-');
            return string.IsNullOrWhiteSpace(stem) ? SlugifyToFileName(pageTitle) : stem;
        }

        return SlugifyToFileName(pageTitle);
    }

    private static string SlugifyToFileName(string title)
    {
        var s = title.Trim();
        s = Regex.Replace(s, @"\s+", "-");
        s = Regex.Replace(s, @"[^A-Za-z0-9\-]", "");
        s = Regex.Replace(s, @"\-{2,}", "-");
        return s.Trim('-');
    }

    // -------------------------
    // Layout parsing helpers
    // -------------------------

    private static (string? sectionId, string? columnId) ExtractFirstSectionAndColumnIds(string pageJson)
    {
        using var doc = JsonDocument.Parse(pageJson);
        var root = doc.RootElement;

        if (!root.TryGetProperty("canvasLayout", out var canvasLayout)) return (null, null);
        if (!canvasLayout.TryGetProperty("horizontalSections", out var sections)) return (null, null);
        if (sections.ValueKind != JsonValueKind.Array || sections.GetArrayLength() == 0) return (null, null);

        var firstSection = sections[0];
        var sectionId = firstSection.TryGetProperty("id", out var sId) ? sId.GetString() : null;

        if (!firstSection.TryGetProperty("columns", out var columns)) return (sectionId, null);
        if (columns.ValueKind != JsonValueKind.Array || columns.GetArrayLength() == 0) return (sectionId, null);

        var firstColumn = columns[0];
        var columnId = firstColumn.TryGetProperty("id", out var cId) ? cId.GetString() : null;

        return (sectionId, columnId);
    }

    // -------------------------
    // Response helpers
    // -------------------------

    private static async Task<HttpResponseData> Json(HttpRequestData req, HttpStatusCode status, string json)
    {
        var resp = req.CreateResponse(status);
        resp.Headers.Add("Content-Type", "application/json");
        await resp.WriteStringAsync(json);
        return resp;
    }

    private static async Task<HttpResponseData> Text(HttpRequestData req, HttpStatusCode status, string text)
    {
        var resp = req.CreateResponse(status);
        resp.Headers.Add("Content-Type", "text/plain; charset=utf-8");
        await resp.WriteStringAsync(text);
        return resp;
    }

    private static async Task<HttpResponseData> Error(
        HttpRequestData req,
        HttpStatusCode status,
        string error,
        string message,
        string correlationId)
    {
        var payload = new
        {
            error,
            message,
            correlationId
        };

        return await Json(req, status, JsonSerializer.Serialize(payload));
    }

    private static string TryGetServerRelativeUrl(string? pageUrl)
    {
        if (string.IsNullOrWhiteSpace(pageUrl))
            return string.Empty;

        return Uri.TryCreate(pageUrl, UriKind.Absolute, out var uri)
            ? uri.AbsolutePath
            : string.Empty;
    }

    private static async Task<(bool success, string detail)> TryCheckInDraftAsync(string siteUrl, string pageWebUrl, ILogger log)
    {
        try
        {
            var siteUri = new Uri(siteUrl);
            var pageUri = new Uri(pageWebUrl);

            var serverRelativePath = pageUri.AbsolutePath;
            var escapedPath = serverRelativePath.Replace("'", "''");

            var checkInUrl =
                $"{siteUri.Scheme}://{siteUri.Host}{siteUri.AbsolutePath.TrimEnd('/')}/_api/web/GetFileByServerRelativePath(decodedurl='{escapedPath}')/CheckIn(comment='AI draft check-in',checkintype=0)";

            var credential = new DefaultAzureCredential();
            var token = await credential.GetTokenAsync(
                new TokenRequestContext(new[] { $"{siteUri.Scheme}://{siteUri.Host}/.default" }));

            using var request = new HttpRequestMessage(HttpMethod.Post, checkInUrl);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            using var response = await Http.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                log.LogWarning(
                    "Draft check-in failed. Status={StatusCode}, Host={Host}, BodyLength={BodyLength}",
                    (int)response.StatusCode,
                    siteUri.Host,
                    body?.Length ?? 0);
                return (false, $"status={(int)response.StatusCode};host={siteUri.Host};bodyLength={(body?.Length ?? 0)}");
            }

            log.LogInformation("Draft check-in succeeded for {ServerRelativePath}", serverRelativePath);
            return (true, "ok");
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Draft check-in failed with exception.");
            return (false, ex.Message);
        }
    }

    private static string TruncateHeaderValue(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
            return value;

        return value.Substring(0, maxLength);
    }

    private static string SafeHostFromUrl(string url)
    {
        if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return uri.Host;

        return "unknown";
    }

    private static (string? pageId, string? webUrl) TryExtractPageIdAndWebUrl(string pageJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(pageJson);
            var root = doc.RootElement;

            var pageId = root.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            var webUrl = root.TryGetProperty("webUrl", out var webUrlEl) ? webUrlEl.GetString() : null;

            return (pageId, webUrl);
        }
        catch
        {
            return (null, null);
        }
    }

    // -------------------------
    // Request model
    // -------------------------

    private class CreateKbDraftRequest
    {
        [JsonPropertyName("fileName")]
        public string? FileName { get; set; } = null;

        [JsonPropertyName("siteUrl")]
        public string SiteUrl { get; set; } = "";

        [JsonPropertyName("templateServerRelativePath")]
        public string TemplateServerRelativePath { get; set; } = "";

        [JsonPropertyName("pageTitle")]
        public string PageTitle { get; set; } = "";

        [JsonPropertyName("htmlContent")]
        public string HtmlContent { get; set; } = "";

        [JsonPropertyName("checkInDraft")]
        public bool CheckInDraft { get; set; } = true;
    }

    private sealed class GraphApiException : Exception
    {
        public HttpStatusCode StatusCode { get; }
        public string Method { get; }
        public string Url { get; }
        public string? GraphCode { get; }
        public string? GraphMessage { get; }
        public string ResponseBody { get; }
        public string? RequestId { get; }
        public string? ClientRequestId { get; }
        public int? RetryAfterSeconds { get; }
        public int Attempt { get; }
        public int MaxAttempts { get; }

        public GraphApiException(
            string message,
            HttpStatusCode statusCode,
            string method,
            string url,
            string? graphCode,
            string? graphMessage,
            string responseBody,
            string? requestId,
            string? clientRequestId,
            int? retryAfterSeconds,
            int attempt,
            int maxAttempts)
            : base(message)
        {
            StatusCode = statusCode;
            Method = method;
            Url = url;
            GraphCode = graphCode;
            GraphMessage = graphMessage;
            ResponseBody = responseBody;
            RequestId = requestId;
            ClientRequestId = clientRequestId;
            RetryAfterSeconds = retryAfterSeconds;
            Attempt = attempt;
            MaxAttempts = maxAttempts;
        }
    }
}
