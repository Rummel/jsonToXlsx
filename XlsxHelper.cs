using System.Text.Json;

internal static class XlsxHelper
{
    public static XlsxData? LoadJsonFile(string fileName)
    {
        var jsonString = File.ReadAllText(fileName);
        var options = new JsonSerializerOptions
        {
            AllowTrailingCommas = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        var data = JsonSerializer.Deserialize<XlsxData>(jsonString, options);
        return data;
    }

    public static object JsonElementToNumber(JsonElement item)
    {
        if (item.TryGetInt32(out var i))
        {
            return i;
        }
        if (item.TryGetDouble(out var d))
        {
            return d;
        }
        if (item.TryGetDecimal(out var dec))
        {
            return dec;
        }
        return item.ToString();
    }
}