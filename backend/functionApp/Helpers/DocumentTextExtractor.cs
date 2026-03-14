using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace functionApp.Helpers;

/// <summary>
/// Extracts text content from various document file types (.pdf, .docx, .doc, and plain text formats).
/// </summary>
public static class DocumentTextExtractor
{
    /// <summary>
    /// Extracts text content from file bytes based on file type.
    /// Supports .txt, .csv, .json, .xml, .html, .md, .pdf, .docx, and .doc files.
    /// Returns null if content is empty; returns a placeholder for unsupported types.
    /// </summary>
    public static string? ExtractTextContent(byte[]? fileContent, string? fileName, int maxLength = 5000)
    {
        if (fileContent == null || fileContent.Length == 0)
            return null;

        var extension = fileName != null ? Path.GetExtension(fileName) : null;

        try
        {
            var text = extension?.ToLowerInvariant() switch
            {
                ".pdf" => ExtractTextFromPdf(fileContent),
                ".docx" => ExtractTextFromDocx(fileContent),
                ".doc" => ExtractTextFromDoc(fileContent),
                ".txt" or ".csv" or ".json" or ".xml" or ".html" or ".htm" or ".md" =>
                    Encoding.UTF8.GetString(fileContent),
                _ => null
            };

            if (text == null)
                return $"[Unsupported file type: {fileName}, size: {fileContent.Length} bytes]";

            text = text.Trim();
            if (string.IsNullOrEmpty(text))
                return $"[No text content extracted from: {fileName}]";

            if (text.Length > maxLength)
                text = text[..maxLength] + $"\n... [truncated, total {text.Length} characters]";

            return text;
        }
        catch
        {
            return $"[Could not extract text from: {fileName}, size: {fileContent.Length} bytes]";
        }
    }

    /// <summary>
    /// Native PDF text extraction — parses stream objects, decompresses deflate streams,
    /// and extracts text from BT/ET blocks using Tj and TJ operators.
    /// </summary>
    private static string ExtractTextFromPdf(byte[] pdfBytes)
    {
        // Decode bytes as Latin1 to preserve raw byte values for stream parsing
        var pdfText = Encoding.Latin1.GetString(pdfBytes);

        // Find all stream...endstream blocks
        var streamRegex = new Regex(@"stream\r?\n([\s\S]*?)endstream", RegexOptions.Compiled);
        var streams = new List<byte[]>();

        foreach (Match match in streamRegex.Matches(pdfText))
        {
            var streamStr = match.Groups[1].Value;
            var streamBytes = new byte[streamStr.Length];
            for (int i = 0; i < streamStr.Length; i++)
            {
                streamBytes[i] = (byte)(streamStr[i] & 0xFF);
            }
            streams.Add(streamBytes);
        }

        var allTextParts = new List<string>();

        foreach (var streamBytes in streams)
        {
            try
            {
                byte[] contentBytes = streamBytes;

                // Check if content starts with ASCII85 encoded data
                var startStr = Encoding.Latin1.GetString(contentBytes, 0, Math.Min(4, contentBytes.Length));
                if (Regex.IsMatch(startStr, @"^[!-u~]"))
                {
                    contentBytes = DecodeAscii85(contentBytes);
                }

                // Try to decompress (most PDF streams are Flate/deflate compressed)
                byte[] decompressed;
                try
                {
                    decompressed = DeflateDecompress(contentBytes);
                }
                catch
                {
                    decompressed = contentBytes;
                }

                var content = Encoding.Latin1.GetString(decompressed);

                // Find BT...ET text blocks
                var btBlocks = Regex.Matches(content, @"BT[\s\S]*?ET");
                foreach (Match block in btBlocks)
                {
                    var blockText = block.Value;

                    // Tj operator: (text) Tj
                    var tjRegex = new Regex(@"\(([^)]*)\)\s*(?:Tj|'|"")");
                    foreach (Match tjMatch in tjRegex.Matches(blockText))
                    {
                        var decoded = DecodePdfString(tjMatch.Groups[1].Value);
                        if (!string.IsNullOrWhiteSpace(decoded))
                            allTextParts.Add(decoded.Trim());
                    }

                    // TJ array operator: [(text) num (text)] TJ
                    var tjArrRegex = new Regex(@"\[([^\]]*)\]\s*TJ");
                    foreach (Match tjArrMatch in tjArrRegex.Matches(blockText))
                    {
                        var innerRegex = new Regex(@"\(([^)]*)\)");
                        foreach (Match innerMatch in innerRegex.Matches(tjArrMatch.Groups[1].Value))
                        {
                            var decoded = DecodePdfString(innerMatch.Groups[1].Value);
                            if (!string.IsNullOrWhiteSpace(decoded))
                                allTextParts.Add(decoded.Trim());
                        }
                    }
                }
            }
            catch
            {
                // Skip streams that can't be processed
            }
        }

        return Regex.Replace(
            Regex.Replace(
                string.Join(" ", allTextParts),
                @"[ \t]{2,}", " "),
            @"\n{3,}", "\n\n").Trim();
    }

    /// <summary>
    /// Decodes an ASCII85 (base85) encoded byte array.
    /// </summary>
    private static byte[] DecodeAscii85(byte[] data)
    {
        var input = Encoding.Latin1.GetString(data).Trim();

        // Strip <~ and ~> delimiters if present
        if (input.StartsWith("<~"))
            input = input[2..];
        if (input.EndsWith("~>"))
            input = input[..^2];

        var result = new List<byte>();
        int i = 0;

        while (i < input.Length)
        {
            if (input[i] == 'z')
            {
                // 'z' is shorthand for four zero bytes
                result.AddRange(new byte[4]);
                i++;
                continue;
            }

            // Collect up to 5 characters
            var group = new List<int>();
            while (group.Count < 5 && i < input.Length)
            {
                if (input[i] >= '!' && input[i] <= 'u')
                {
                    group.Add(input[i] - 33);
                }
                i++;
            }

            if (group.Count == 0) break;

            // Pad with 'u' (84) values if fewer than 5
            int padding = 5 - group.Count;
            for (int p = 0; p < padding; p++)
                group.Add(84);

            long value = 0;
            for (int j = 0; j < 5; j++)
                value = value * 85 + group[j];

            var bytes = new[]
            {
                (byte)((value >> 24) & 0xFF),
                (byte)((value >> 16) & 0xFF),
                (byte)((value >> 8) & 0xFF),
                (byte)(value & 0xFF)
            };

            // Only add the non-padded bytes
            for (int j = 0; j < 4 - padding; j++)
                result.Add(bytes[j]);
        }

        return result.ToArray();
    }

    /// <summary>
    /// Deflate-decompress a byte array (skips 2-byte zlib header if present).
    /// </summary>
    private static byte[] DeflateDecompress(byte[] data)
    {
        // Check for zlib header (0x78 0x01/9C/DA) and skip it
        int offset = 0;
        if (data.Length >= 2 && data[0] == 0x78)
            offset = 2;

        using var input = new MemoryStream(data, offset, data.Length - offset);
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        deflate.CopyTo(output);
        return output.ToArray();
    }

    /// <summary>
    /// Decodes PDF escape sequences in a string (e.g. \n, \r, octal codes).
    /// </summary>
    private static string DecodePdfString(string input)
    {
        var sb = new StringBuilder(input.Length);
        for (int i = 0; i < input.Length; i++)
        {
            if (input[i] == '\\' && i + 1 < input.Length)
            {
                i++;
                switch (input[i])
                {
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case '(': sb.Append('('); break;
                    case ')': sb.Append(')'); break;
                    case '\\': sb.Append('\\'); break;
                    default:
                        // Octal escape (e.g. \012)
                        if (input[i] >= '0' && input[i] <= '7')
                        {
                            var octal = input[i].ToString();
                            if (i + 1 < input.Length && input[i + 1] >= '0' && input[i + 1] <= '7')
                            {
                                octal += input[++i];
                                if (i + 1 < input.Length && input[i + 1] >= '0' && input[i + 1] <= '7')
                                    octal += input[++i];
                            }
                            sb.Append((char)Convert.ToInt32(octal, 8));
                        }
                        else
                        {
                            sb.Append(input[i]);
                        }
                        break;
                }
            }
            else
            {
                sb.Append(input[i]);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Extracts text from a .docx file (Office Open XML).
    /// The .docx format is a ZIP archive containing word/document.xml with the document body.
    /// </summary>
    private static string ExtractTextFromDocx(byte[] docxBytes)
    {
        using var stream = new MemoryStream(docxBytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var documentEntry = archive.GetEntry("word/document.xml");
        if (documentEntry == null)
            return string.Empty;

        using var entryStream = documentEntry.Open();
        var xDoc = XDocument.Load(entryStream);

        // Word namespace for the main document content
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        var paragraphs = xDoc.Descendants(w + "p");
        var sb = new StringBuilder();

        foreach (var para in paragraphs)
        {
            var runs = para.Descendants(w + "t");
            foreach (var run in runs)
            {
                sb.Append(run.Value);
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// Extracts text from an older .doc file (binary OLE2 format).
    /// Scans for runs of printable ASCII/Unicode characters in UTF-16LE encoding.
    /// </summary>
    private static string ExtractTextFromDoc(byte[] docBytes)
    {
        var sb = new StringBuilder();
        var currentRun = new StringBuilder();

        // Word stores body text as UTF-16LE in the binary
        if (docBytes.Length >= 2)
        {
            for (int i = 0; i < docBytes.Length - 1; i += 2)
            {
                char c = (char)(docBytes[i] | (docBytes[i + 1] << 8));

                if (c == '\r' || c == '\n' || c == '\t' || (c >= ' ' && c < 0x7F))
                {
                    currentRun.Append(c);
                }
                else
                {
                    if (currentRun.Length >= 10) // Only keep runs of 10+ chars to filter noise
                    {
                        sb.Append(currentRun);
                        sb.Append(' ');
                    }
                    currentRun.Clear();
                }
            }

            if (currentRun.Length >= 10)
                sb.Append(currentRun);
        }

        var result = sb.ToString().Trim();

        // Collapse excessive whitespace
        result = Regex.Replace(result, @"[ \t]{2,}", " ");
        result = Regex.Replace(result, @"\n{3,}", "\n\n");

        return result;
    }
}
