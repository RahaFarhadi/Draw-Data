using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using netDxf;
/// <summary>
/// Represents a single column mapping entry coming from mapping.json.
/// </summary>
class MappingColumn
{
    public string col { get; set; } = string.Empty;
    public string? attr { get; set; }
    public string? source { get; set; }
}

/// <summary>
/// Root object for mapping.json.
/// </summary>
class Mapping
{
    public List<MappingColumn> columns { get; set; } = new();
}
class Program
{
    private static readonly HashSet<string> SingleCharColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "R", "B", "Y", "L", "W", "G", "V", "O", "P", "S", "T", "BR", "GR", "PI", "LB", "VI", "GY"
    };

    private static readonly string[] WireTypePatterns =
    {
        "AVSS", "AVS", "FLRY", "T1", "T2", "T3", "T4", "GPT", "CAVS"
    };
    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        var (dxfPath, mappingPath) = ResolvePaths(args);
        if (!ValidateInputFiles(dxfPath, mappingPath)) return;

        var mapping = LoadMapping(mappingPath);
        if (mapping.columns.Count == 0)
        {
            Console.WriteLine("ERROR: mapping.json does not contain any columns.");
            return;
        }

        var dxf = LoadDxf(dxfPath);
        if (dxf is null) return;

        var table = CreateWireTable(mapping);

        if (!TryExtractFromInserts(dxf, table, mapping))
        {
            Console.WriteLine("⚠️ No structured attribute data found. Falling back to parsing Text/MText entities.");
            ParseTextEntities(dxf, table);
        }

        SaveToExcel(table, Path.Combine(Path.GetDirectoryName(dxfPath) ?? ".", "output.xlsx"));
    }
    private static (string dxfPath, string mappingPath) ResolvePaths(string[] args)
    {
        if (args.Length >= 2)
        {
            return (args[0], args[1]);
        }

        // Fallback to historical defaults for backward compatibility.
        var dxfPath = Environment.GetEnvironmentVariable("DXF_PATH") ?? "C:/tmp/sample-dxf.dxf";
        var mappingPath = Environment.GetEnvironmentVariable("MAPPING_PATH") ??
                          "C:/Users/Digibod.ir/source/repos/dxf-ext/dxf-ext/mapping.json";
        return (dxfPath, mappingPath);
    }

    private static bool ValidateInputFiles(string dxfPath, string mappingPath)
    {
        if (!File.Exists(dxfPath))
        {
            Console.WriteLine($"ERROR: DXF file not found: {dxfPath}");
            return false;
        }

        if (!File.Exists(mappingPath))
        {
            Console.WriteLine($"ERROR: mapping.json not found: {mappingPath}");
            return false;
        }

        return true;
    }

    // Load DXF
    private static Mapping LoadMapping(string mappingPath)
    {
        var json = File.ReadAllText(mappingPath);
        var mapping = JsonSerializer.Deserialize<Mapping>(json) ?? new Mapping();
        mapping.columns = mapping.columns
            .Where(c => !string.IsNullOrWhiteSpace(c.col))
            .ToList();
        return mapping;
    }

    private static DxfDocument? LoadDxf(string dxfPath)
    {
        try
        {
            var document = DxfDocument.Load(dxfPath);
            Console.WriteLine($"✅ DXF loaded. Entities count: {document.Entities.All.Count()}");
            return document;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR loading DXF: {ex.Message}");
            return null;
        }
    }

    // Utility: get enumerable from a possible property or method using reflection
    private static DataTable CreateWireTable(Mapping mapping)
    {
        var table = new DataTable("Wires");
        foreach (var column in mapping.columns)
        {
            if (!table.Columns.Contains(column.col))
            {
                table.Columns.Add(column.col);
            }
        }
        return table;
    }

    private static bool TryExtractFromInserts(DxfDocument dxf, DataTable table, Mapping mapping)
    {
        var inserts = dxf.Entities.All.OfType<netDxf.Entities.Insert>().ToList();
        if (!inserts.Any()) return false;

        var indexColumn = mapping.columns.FirstOrDefault(c => c.source == "auto_index")?.col;
        var targetColumns = mapping.columns.Where(c => !string.IsNullOrEmpty(c.attr)).ToList();
        if (!targetColumns.Any()) return false;

        var addedRows = 0;
        foreach (var insert in inserts)
        {
            var attributes = insert.Attributes.ToList();
            if (attributes.Count == 0) continue;

            var row = table.NewRow();
            foreach (DataColumn column in table.Columns) row[column.ColumnName] = string.Empty;

            if (!string.IsNullOrEmpty(indexColumn))
            {
                row[indexColumn] = (addedRows + 1).ToString(CultureInfo.InvariantCulture);
            }

            foreach (var column in targetColumns)
            {
                var attribute = attributes.FirstOrDefault(a =>
                    string.Equals(a.Tag, column.attr, StringComparison.OrdinalIgnoreCase));
                if (attribute != null)
                {
                    row[column.col] = attribute.Value;
                }
            }

            table.Rows.Add(row);
            addedRows++;
        }

        if (addedRows > 0)
        {
            Console.WriteLine($"✅ Successfully extracted data from {addedRows} Insert entities.");
            return true;
        }

        return false;
    }

    private static void ParseTextEntities(DxfDocument dxf, DataTable table)
    {
        var texts = dxf.Entities.All
            .Where(e => e is netDxf.Entities.Text || e is netDxf.Entities.MText)
            .Select(entity =>
            {
                if (entity is netDxf.Entities.Text text)
                {
                    return (text.Value, text.Position.X, text.Position.Y);
                }
                var mtext = (netDxf.Entities.MText)entity;
                return (mtext.Value, mtext.Position.X, mtext.Position.Y);
            })
            .Select(item => (value: CleanText(item.Item1), item.Item2, item.Item3))
            .Where(item => !string.IsNullOrWhiteSpace(item.value) && !IsNoise(item.value))
            .ToList();

        Console.WriteLine($"Collected {texts.Count} cleaned text-like entities.");

        var groupedRows = GroupByRow(texts);
        Console.WriteLine($"Grouped into {groupedRows.Count} rows based on Y coordinate.");

        var idx = 1;
        foreach (var group in groupedRows)
        {
            var row = table.NewRow();
            foreach (DataColumn column in table.Columns) row[column.ColumnName] = string.Empty;

            if (table.Columns.Contains("رديف"))
            {
                row["رديف"] = idx.ToString(CultureInfo.InvariantCulture);
            }
            PopulateRowFromGroup(table, group, row);

            table.Rows.Add(row);
            idx++;
        }
    }
    private static void PopulateRowFromGroup(DataTable table, List<(string value, double x, double y)> group, DataRow row)
    {
        var sortedGroup = group.OrderBy(item => item.x).ToList();
        var rowData = new Dictionary<string, string>();
        var usedValues = new HashSet<string>();


        // Stage 1: absolute identifiers
        foreach (var item in sortedGroup)
        {
            var value = item.value.Trim().ToUpperInvariant();
            if (usedValues.Contains(value)) continue;

            if (table.Columns.Contains("رنگ سيم") && !rowData.ContainsKey("رنگ سيم") &&
                SingleCharColors.Contains(value))
            {
                rowData["رنگ سيم"] = value;
                usedValues.Add(value);
                continue;
            }
            if (!rowData.ContainsKey("نوع سيم"))
            {
                var matchedPattern = WireTypePatterns
                    .OrderByDescending(p => p.Length)
                    .FirstOrDefault(pattern => value.EndsWith(pattern, StringComparison.OrdinalIgnoreCase));

                if (matchedPattern != null)
                {
                    var sizePart = value[..^matchedPattern.Length].Trim();
                    var cleanValue = sizePart.Replace(',', '.').Replace(" ", string.Empty);

                    if (double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) &&
                        size is >= 0.1 and <= 5.0)
                    {
                        if (table.Columns.Contains("سايزسيم"))
                        {
                            rowData["سايزسيم"] = sizePart;
                        }
                        rowData["نوع سيم"] = matchedPattern;
                        usedValues.Add(value);
                        continue;
                    }

                    if (string.IsNullOrEmpty(sizePart))
                    {
                        rowData["نوع سيم"] = matchedPattern;
                        usedValues.Add(value);
                        continue;
                    }
                }
            }
        }

        // Stage 2: numbers (length / size)
        foreach (var item in sortedGroup)
        {
            var rawValue = item.value.Trim().ToUpperInvariant();
            if (usedValues.Contains(rawValue)) continue;

            if (int.TryParse(rawValue, out var length))
            {
                if (table.Columns.Contains("طول برش سيم") && !rowData.ContainsKey("طول برش سيم") &&
                    length is >= 50 and <= 2500)
                {
                    rowData["طول برش سيم"] = rawValue;
                    usedValues.Add(rawValue);
                }
                continue;
            }

            var cleanValue = rawValue.Replace(',', '.').Replace(" ", string.Empty);
            if (table.Columns.Contains("سايزسيم") && !rowData.ContainsKey("سايزسيم") &&
                double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) &&
                size is >= 0.1 and <= 5.0 && cleanValue.Length <= 5)
            {
                rowData["سايزسيم"] = rawValue;
                usedValues.Add(rawValue);
            }
        }

        // Stage 3: wire code and connectors
        foreach (var item in sortedGroup)
        {
            var value = item.value.Trim().ToUpperInvariant();
            if (usedValues.Contains(value)) continue;

            bool looksLikeCode = value.Length is >= 3 and <= 10 && value.Any(char.IsLetter);
            bool looksLikeConnector = value.Length is >= 4 and <= 15 && value.Any(char.IsLetter);
            bool isNoise = IsNoise(value) || SingleCharColors.Contains(value) ||
                           WireTypePatterns.Any(p => value.Equals(p, StringComparison.OrdinalIgnoreCase));

            if (looksLikeCode && !isNoise && table.Columns.Contains("کدسیم") && !rowData.ContainsKey("کدسیم"))
            {
                rowData["کدسیم"] = value;
                usedValues.Add(value);
                continue;
            }

            if (looksLikeConnector && !isNoise && !value.All(char.IsDigit))
            {
                if (table.Columns.Contains("ابتدا") && !rowData.ContainsKey("ابتدا"))
                {
                    rowData["ابتدا"] = value;
                    usedValues.Add(value);
                    continue;
                }

                if (table.Columns.Contains("انتها") && !rowData.ContainsKey("انتها"))
                {
                    rowData["انتها"] = value;
                    usedValues.Add(value);
                }
            }
        }

        foreach (var kvp in rowData)
        {
            row[kvp.Key] = kvp.Value;
        }
    }

    private static string CleanText(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var cleaned = Regex.Replace(raw, @"\\(P|H|W|S|T|L|O|U|A|C|F)[^;]*;?", string.Empty);
        cleaned = cleaned.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        return cleaned.Trim();
    }

    private static bool IsNoise(string value)
    {
        var v = value.ToUpperInvariant();
        return v.StartsWith("PIN", StringComparison.Ordinal) ||
               v.EndsWith(":", StringComparison.Ordinal) ||
               v.Contains("SEAL", StringComparison.Ordinal) ||
               v.Contains("CLIP", StringComparison.Ordinal) ||
               v.Contains("AMP", StringComparison.Ordinal) ||
               v.Contains("YAZAKI", StringComparison.Ordinal) ||
               v.Contains("KET", StringComparison.Ordinal) ||
               v.Contains("KUM", StringComparison.Ordinal) ||
               v.Contains("SWS", StringComparison.Ordinal) ||
               v.Contains("NOTE", StringComparison.Ordinal) ||
               v.Contains("SPECIFICATION", StringComparison.Ordinal) ||
               v.Contains("ASSY", StringComparison.Ordinal) ||
               v.Contains("DESCRIPTION", StringComparison.Ordinal);
    }

    private static List<List<(string value, double x, double y)>> GroupByRow(List<(string value, double x, double y)> texts)
    {
        const double yTolerance = 5.0;
        var groups = new List<List<(string value, double x, double y)>>();

        foreach (var text in texts.OrderByDescending(t => t.y).ThenBy(t => t.x))
        {
            var group = groups.FirstOrDefault(gr => gr.Any() && Math.Abs(gr.Average(i => i.y) - text.y) <= yTolerance);
            if (group is null)
            {
                groups.Add(new List<(string, double, double)> { text });
            }
            else
            {
                group.Add(text);
            }
        }

        return groups;
    }

    private static void SaveToExcel(DataTable table, string outputPath)
    {
        try
        {
            using var workbook = new XLWorkbook();
            workbook.Worksheets.Add(table, "Wires");
            workbook.Worksheet(1).Columns().AdjustToContents();
            workbook.SaveAs(outputPath);
            Console.WriteLine($"✅ Excel saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR saving Excel: {ex.Message}");
        }
    }
}
