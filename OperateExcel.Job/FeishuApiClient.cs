using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Options;

namespace OperateExcel.Job;

public sealed class FeishuApiClient
{
    private const string ApiBaseUrl = "https://open.feishu.cn/open-apis";
    private const string AttachmentReadFailed = "\u9644\u4ef6\u8bfb\u53d6\u5931\u8d25";

    private readonly HttpClient _httpClient;
    private readonly FeishuOptions _options;
    private readonly ILogger<FeishuApiClient> _logger;
    private string? _tenantAccessToken;

    public FeishuApiClient(HttpClient httpClient, IOptions<FeishuOptions> options, ILogger<FeishuApiClient> logger)
    {
        _httpClient = httpClient;
        _options = options.Value;
        _logger = logger;
    }

    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(_options.AppId)
        && !string.IsNullOrWhiteSpace(_options.AppSecret)
        && (!string.IsNullOrWhiteSpace(_options.BitableAppToken) || TryReadWikiNodeToken(_options.WikiUrl, out _))
        && !string.IsNullOrWhiteSpace(ResolveTableId());

    public async Task<byte[]> DownloadProfitAttachmentAsync(DateOnly processingDate, CancellationToken cancellationToken = default)
    {
        var accessToken = await GetTenantAccessTokenAsync(cancellationToken);
        var appToken = await ResolveBitableAppTokenAsync(accessToken, cancellationToken);
        var tableId = ResolveTableId()
            ?? throw new InvalidOperationException("Feishu table id is not configured.");

        var fileToken = await GetProfitAttachmentFileTokenAsync(
            accessToken,
            appToken,
            tableId,
            processingDate,
            cancellationToken);

        var downloadUrl = await GetTemporaryDownloadUrlAsync(accessToken, fileToken, cancellationToken);
        using var request = new HttpRequestMessage(HttpMethod.Get, downloadUrl);
        using var response = await _httpClient.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsByteArrayAsync(cancellationToken);
    }

    private async Task<string> GetTenantAccessTokenAsync(CancellationToken cancellationToken)
    {
        if (_tenantAccessToken is not null)
        {
            return _tenantAccessToken;
        }

        using var request = CreateJsonPost(
            $"{ApiBaseUrl}/auth/v3/tenant_access_token/internal",
            new
            {
                app_id = _options.AppId,
                app_secret = _options.AppSecret
            });

        using var response = await _httpClient.SendAsync(request, cancellationToken);
        using var json = await ReadSuccessJsonAsync(response, "Get tenant_access_token failed", cancellationToken);
        _tenantAccessToken = ReadRequiredString(json.RootElement, "tenant_access_token", "tenant_access_token");
        return _tenantAccessToken;
    }

    private async Task<string> ResolveBitableAppTokenAsync(string accessToken, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(_options.BitableAppToken))
        {
            return _options.BitableAppToken;
        }

        if (!TryReadWikiNodeToken(_options.WikiUrl, out var wikiNodeToken))
        {
            throw new InvalidOperationException("Feishu wiki node token is not configured.");
        }

        using var request = CreateAuthorizedRequest(
            HttpMethod.Get,
            $"{ApiBaseUrl}/wiki/v2/spaces/get_node?token={Uri.EscapeDataString(wikiNodeToken)}",
            accessToken);

        using var response = await _httpClient.SendAsync(request, cancellationToken);
        using var json = await ReadSuccessJsonAsync(response, "Get Feishu wiki node failed", cancellationToken);
        var node = json.RootElement.GetProperty("data").GetProperty("node");
        return ReadRequiredString(node, "obj_token", "wiki node obj_token");
    }

    private async Task<string> GetProfitAttachmentFileTokenAsync(
        string accessToken,
        string appToken,
        string tableId,
        DateOnly processingDate,
        CancellationToken cancellationToken)
    {
        var dateFilter = new
        {
            conjunction = "and",
            conditions = new object[]
            {
                new
                {
                    field_name = _options.DateFieldName,
                    @operator = "is",
                    value = new[]
                    {
                        "ExactDate",
                        GetLocalDateStartUnixMilliseconds(processingDate).ToString(CultureInfo.InvariantCulture)
                    }
                }
            }
        };

        var records = await SearchRecordsAsync(accessToken, appToken, tableId, dateFilter, cancellationToken);
        if (records.Count == 0)
        {
            records = await SearchRecordsAsync(accessToken, appToken, tableId, filter: null, cancellationToken);
            records = records
                .Where(record => record.TryGetProperty(_options.DateFieldName, out var dateField)
                    && IsSameLocalDate(dateField, processingDate))
                .ToList();
        }

        if (records.Count == 0)
        {
            throw CreateAttachmentFailure($"No Feishu record found for date {processingDate:yyyy-MM-dd}.");
        }

        if (records.Count > 1)
        {
            _logger.LogWarning(
                "Feishu record query returned {RecordCount} records for {ProcessingDate}; using the first record.",
                records.Count,
                processingDate);
        }

        var fields = records[0];
        if (!fields.TryGetProperty(_options.AttachmentFieldName, out var attachmentField)
            || attachmentField.ValueKind != JsonValueKind.Array)
        {
            throw CreateAttachmentFailure($"Field {_options.AttachmentFieldName} has no attachment.");
        }

        if (attachmentField.GetArrayLength() != 1)
        {
            throw CreateAttachmentFailure($"Field {_options.AttachmentFieldName} attachment count is {attachmentField.GetArrayLength()}.");
        }

        var attachment = attachmentField[0];
        var fileName = TryReadString(attachment, "name") ?? TryReadString(attachment, "file_name") ?? string.Empty;
        if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw CreateAttachmentFailure($"Attachment is not an .xlsx file: {fileName}");
        }

        return TryReadString(attachment, "file_token")
            ?? TryReadString(attachment, "token")
            ?? throw CreateAttachmentFailure("Attachment is missing file_token.");
    }

    private async Task<List<JsonElement>> SearchRecordsAsync(
        string accessToken,
        string appToken,
        string tableId,
        object? filter,
        CancellationToken cancellationToken)
    {
        var records = new List<JsonElement>();
        string? pageToken = null;

        do
        {
            var query = new List<string> { "page_size=500" };
            var viewId = ResolveViewId();
            if (!string.IsNullOrWhiteSpace(viewId))
            {
                query.Add($"view_id={Uri.EscapeDataString(viewId)}");
            }

            if (!string.IsNullOrWhiteSpace(pageToken))
            {
                query.Add($"page_token={Uri.EscapeDataString(pageToken)}");
            }

            var body = new Dictionary<string, object>
            {
                ["field_names"] = new[] { _options.DateFieldName, _options.AttachmentFieldName }
            };
            if (filter is not null)
            {
                body["filter"] = filter;
            }

            using var request = CreateAuthorizedJsonPost(
                $"{ApiBaseUrl}/bitable/v1/apps/{Uri.EscapeDataString(appToken)}/tables/{Uri.EscapeDataString(tableId)}/records/search?{string.Join('&', query)}",
                accessToken,
                body);

            using var response = await _httpClient.SendAsync(request, cancellationToken);
            JsonDocument json;
            try
            {
                json = await ReadSuccessJsonAsync(response, "Search Feishu bitable records failed", cancellationToken);
            }
            catch (FeishuApiException exception) when (filter is not null && exception.Message.Contains("InvalidFilter", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning(exception, "Feishu date filter was rejected. Retrying record query without server-side filter.");
                return [];
            }

            using (json)
            {
                var data = json.RootElement.GetProperty("data");
                if (data.TryGetProperty("items", out var items) && items.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        if (item.TryGetProperty("fields", out var fields) && fields.ValueKind == JsonValueKind.Object)
                        {
                            records.Add(fields.Clone());
                        }
                    }
                }

                pageToken = data.TryGetProperty("page_token", out var nextPageToken)
                    && nextPageToken.ValueKind == JsonValueKind.String
                    && !string.IsNullOrWhiteSpace(nextPageToken.GetString())
                        ? nextPageToken.GetString()
                        : null;

                var hasMore = data.TryGetProperty("has_more", out var hasMoreElement)
                    && hasMoreElement.ValueKind == JsonValueKind.True;
                if (!hasMore)
                {
                    pageToken = null;
                }
            }
        }
        while (!string.IsNullOrWhiteSpace(pageToken));

        return records;
    }

    private async Task<string> GetTemporaryDownloadUrlAsync(string accessToken, string fileToken, CancellationToken cancellationToken)
    {
        using var request = CreateAuthorizedRequest(
            HttpMethod.Get,
            $"{ApiBaseUrl}/drive/v1/medias/batch_get_tmp_download_url?file_tokens={Uri.EscapeDataString(fileToken)}",
            accessToken);

        using var response = await _httpClient.SendAsync(request, cancellationToken);
        using var json = await ReadSuccessJsonAsync(response, "Get Feishu temporary download URL failed", cancellationToken);
        var data = json.RootElement.GetProperty("data");

        if (TryFindDownloadUrl(data, fileToken, out var downloadUrl))
        {
            return downloadUrl;
        }

        throw CreateAttachmentFailure("Feishu did not return a usable temporary download URL.");
    }

    private static bool TryFindDownloadUrl(JsonElement element, string fileToken, out string downloadUrl)
    {
        downloadUrl = string.Empty;
        if (element.ValueKind == JsonValueKind.Object)
        {
            if (element.TryGetProperty("file_token", out var tokenElement)
                && string.Equals(tokenElement.GetString(), fileToken, StringComparison.Ordinal)
                && TryReadKnownUrlProperty(element, out downloadUrl))
            {
                return true;
            }

            if (element.TryGetProperty(fileToken, out var tokenValue)
                && tokenValue.ValueKind == JsonValueKind.String)
            {
                downloadUrl = tokenValue.GetString() ?? string.Empty;
                return downloadUrl.Length > 0;
            }

            if (TryReadKnownUrlProperty(element, out downloadUrl))
            {
                return true;
            }

            foreach (var property in element.EnumerateObject())
            {
                if (TryFindDownloadUrl(property.Value, fileToken, out downloadUrl))
                {
                    return true;
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                if (TryFindDownloadUrl(item, fileToken, out downloadUrl))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TryReadKnownUrlProperty(JsonElement element, out string downloadUrl)
    {
        foreach (var propertyName in new[] { "tmp_download_url", "download_url", "url" })
        {
            if (element.TryGetProperty(propertyName, out var value)
                && value.ValueKind == JsonValueKind.String
                && !string.IsNullOrWhiteSpace(value.GetString()))
            {
                downloadUrl = value.GetString()!;
                return true;
            }
        }

        downloadUrl = string.Empty;
        return false;
    }

    private async Task<JsonDocument> ReadSuccessJsonAsync(HttpResponseMessage response, string action, CancellationToken cancellationToken)
    {
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new FeishuApiException($"{action}: HTTP {(int)response.StatusCode}, {body}");
        }

        var json = JsonDocument.Parse(body);
        if (json.RootElement.TryGetProperty("code", out var code) && code.GetInt32() != 0)
        {
            var message = TryReadString(json.RootElement, "msg") ?? TryReadString(json.RootElement, "message") ?? body;
            throw new FeishuApiException($"{action}: {message}");
        }

        return json;
    }

    private static HttpRequestMessage CreateAuthorizedRequest(HttpMethod method, string url, string accessToken)
    {
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return request;
    }

    private static HttpRequestMessage CreateAuthorizedJsonPost(string url, string accessToken, object body)
    {
        var request = CreateJsonPost(url, body);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return request;
    }

    private static HttpRequestMessage CreateJsonPost(string url, object body)
    {
        return new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
    }

    private string? ResolveTableId()
    {
        if (!string.IsNullOrWhiteSpace(_options.TableId))
        {
            return _options.TableId;
        }

        return TryReadQueryValue(_options.WikiUrl, "table", out var tableId) ? tableId : null;
    }

    private string? ResolveViewId()
    {
        if (!string.IsNullOrWhiteSpace(_options.ViewId))
        {
            return _options.ViewId;
        }

        return TryReadQueryValue(_options.WikiUrl, "view", out var viewId) ? viewId : null;
    }

    private static bool TryReadWikiNodeToken(string wikiUrl, out string token)
    {
        token = string.Empty;
        if (!Uri.TryCreate(wikiUrl, UriKind.Absolute, out var uri))
        {
            return false;
        }

        var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length < 2 || !string.Equals(segments[^2], "wiki", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        token = segments[^1];
        return token.Length > 0;
    }

    private static bool TryReadQueryValue(string url, string name, out string value)
    {
        value = string.Empty;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return false;
        }

        foreach (var pair in uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = pair.Split('=', 2);
            if (parts.Length == 2 && string.Equals(parts[0], name, StringComparison.OrdinalIgnoreCase))
            {
                value = Uri.UnescapeDataString(parts[1]);
                return value.Length > 0;
            }
        }

        return false;
    }

    private static long GetLocalDateStartUnixMilliseconds(DateOnly date)
    {
        var localDateTime = date.ToDateTime(TimeOnly.MinValue);
        var offset = TimeZoneInfo.Local.GetUtcOffset(localDateTime);
        return new DateTimeOffset(localDateTime, offset).ToUnixTimeMilliseconds();
    }

    private static bool IsSameLocalDate(JsonElement value, DateOnly expectedDate)
    {
        if (TryReadLocalDate(value, out var actualDate))
        {
            return actualDate == expectedDate;
        }

        if (value.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in value.EnumerateArray())
            {
                if (TryReadLocalDate(item, out actualDate) && actualDate == expectedDate)
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TryReadLocalDate(JsonElement value, out DateOnly date)
    {
        date = default;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out var timestamp))
        {
            date = DateOnly.FromDateTime(DateTimeOffset.FromUnixTimeMilliseconds(timestamp).LocalDateTime);
            return true;
        }

        if (value.ValueKind != JsonValueKind.String)
        {
            return false;
        }

        var text = value.GetString()?.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var textTimestamp))
        {
            date = DateOnly.FromDateTime(DateTimeOffset.FromUnixTimeMilliseconds(textTimestamp).LocalDateTime);
            return true;
        }

        return DateOnly.TryParse(text, out date);
    }

    private static string ReadRequiredString(JsonElement element, string propertyName, string displayName)
    {
        return TryReadString(element, propertyName)
            ?? throw new InvalidOperationException($"Feishu response missing {displayName}.");
    }

    private static string? TryReadString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
    }

    private InvalidOperationException CreateAttachmentFailure(string detail)
    {
        _logger.LogError("{AttachmentReadFailed}: {Detail}", AttachmentReadFailed, detail);
        return new InvalidOperationException($"{AttachmentReadFailed}: {detail}");
    }

    private sealed class FeishuApiException : InvalidOperationException
    {
        public FeishuApiException(string message)
            : base(message)
        {
        }
    }
}
