using System.Text;

namespace OperateExcel.Job;

internal static class DelimitedTableReader
{
    public static TableData ReadTxt(string path)
    {
        var lines = File.ReadLines(path, DetectEncoding(path)).ToList();
        return BuildTable(lines, line => line.Split('\t').Select(CleanHeader).ToList());
    }

    public static TableData ReadCsv(string path)
    {
        var lines = File.ReadLines(path, DetectEncoding(path)).ToList();
        return BuildTable(lines, ParseCsvLine);
    }

    private static TableData BuildTable(IReadOnlyList<string> lines, Func<string, List<string>> parseLine)
    {
        // Source exports may include title/blank lines before the table; the first row with two cells is treated as headers.
        var headerRowIndex = FindHeaderRow(lines, parseLine);
        if (headerRowIndex < 0)
        {
            return new TableData([], []);
        }

        var headers = parseLine(lines[headerRowIndex]).Select(CleanHeader).ToList();
        var rows = new List<IReadOnlyList<string>>();

        for (var i = headerRowIndex + 1; i < lines.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            var values = parseLine(lines[i]);
            if (values.Any(value => !string.IsNullOrWhiteSpace(value)))
            {
                rows.Add(values);
            }
        }

        return new TableData(headers, rows);
    }

    private static int FindHeaderRow(IReadOnlyList<string> lines, Func<string, List<string>> parseLine)
    {
        for (var i = 0; i < lines.Count; i++)
        {
            var cells = parseLine(lines[i]).Select(CleanHeader).Where(cell => cell.Length > 0).ToList();
            if (cells.Count >= 2)
            {
                return i;
            }
        }

        return -1;
    }

    private static List<string> ParseCsvLine(string line)
    {
        // Keep CSV parsing local so Amazon-style exports with quoted commas do not require another dependency.
        var values = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (ch == ',' && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        values.Add(current.ToString());
        return values;
    }

    private static Encoding DetectEncoding(string path)
    {
        // Most source files are UTF-8; preserve UTF-8 BOM when present so headers cleanly drop the marker later.
        using var stream = File.OpenRead(path);
        Span<byte> bom = stackalloc byte[3];
        var read = stream.Read(bom);
        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        }

        return Encoding.UTF8;
    }

    public static string CleanHeader(string value)
    {
        return value.Trim().Trim('\uFEFF');
    }
}
