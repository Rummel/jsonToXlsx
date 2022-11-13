using ClosedXML.Excel;
using System.Collections.ObjectModel;
using System.Text.Json;
using System.Text.Json.Serialization;

public class XlsxNumberFormat
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Id { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Format { get; set; }
}

public class XlsxStyle
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Color { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public XlsxNumberFormat? NumberFormat { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public XlsxNumberFormat? DateFormat { get; set; }
}

public class XlsxTableRow : Collection<object>
{
    protected override void InsertItem(int index, object item)
    {
        if (item == null)
        {
            return;
        }

        if (item is JsonElement jItem)
        {
            JsonValueKind valueKind = jItem.ValueKind;

            object? newItem = valueKind switch
            {
                JsonValueKind.String => jItem.ToString(),
                JsonValueKind.Number => XlsxHelper.JsonElementToNumber(jItem),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => item.ToString(),
            };
            if (newItem != null)
            {
                base.InsertItem(index, newItem);
            }
        }
    }
    protected override void SetItem(int index, object item)
    {
        base.SetItem(index, item);
    }
}

public class XlsxCell
{
    public string? Cell { get; set; }
    private Object? _value;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Object? Value
    {
        get { return _value; }
        set
        {
            if (value == null)
            {
                _value = null;
                return;
            }

            if (value is JsonElement jItem)
            {
                object? newValue = jItem.ValueKind switch
                {
                    JsonValueKind.Array => JsonSerializer.Deserialize<List<XlsxTableRow>>(jItem),
                    JsonValueKind.String => jItem.ToString(),
                    JsonValueKind.Number => XlsxHelper.JsonElementToNumber(jItem),
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    JsonValueKind.Null => null,
                    _ => value.ToString(),
                };
                if (newValue != null)
                {
                    _value = newValue;
                }
            }
        }
    }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FormulaA1 { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull),
     JsonConverter(typeof(JsonStringEnumConverter))]
    public XLDataType? DataType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public XlsxStyle? Style { get; set; }
}

public class XlsxWorksheet
{
    public string name { get; set; }
    public IList<XlsxCell> cells { get; set; }

    public XlsxWorksheet(string name)
    {
        this.name = name;
        cells = new List<XlsxCell>();
    }

    public XlsxCell AddCell(string cell, Object value)
    {
        var item = new XlsxCell { Cell = cell, Value = value };
        cells.Add(item);
        return item;
    }
}

public class XlsxData
{
    public string? FileNameSource { get; set; }
    public string? FileNameTarget { get; set; }
    public IList<XlsxWorksheet>? Worksheets { get; set; }
}

