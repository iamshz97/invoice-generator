using System.Diagnostics;
using System.IO.Compression;
using System.Security;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;

// Load appsettings.json (committed, generic) then overlay appsettings.local.json (gitignored, personal)
IConfiguration config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json",       optional: false)
    .AddJsonFile("appsettings.local.json", optional: true)   // overrides appsettings.json when present
    .Build();

string templatePath    = Path.Combine(AppContext.BaseDirectory, config["Invoice:TemplatePath"] ?? "InvoiceTemplate.docx");
string outputFolder    = config["Invoice:OutputFolder"] ?? "output";
string outputFilePrefix = config["Invoice:OutputFilePrefix"] ?? "Invoice";

int anchorInvoiceNo = int.Parse(config["Invoice:InvoiceNumbering:AnchorInvoiceNo"] ?? "122");
int anchorYear      = int.Parse(config["Invoice:InvoiceNumbering:AnchorYear"]      ?? "2025");
int anchorMonth     = int.Parse(config["Invoice:InvoiceNumbering:AnchorMonth"]     ?? "2");

// Resolve output folder relative to the working directory
if (!Path.IsPathRooted(outputFolder))
    outputFolder = Path.Combine(Directory.GetCurrentDirectory(), outputFolder);

Directory.CreateDirectory(outputFolder);

// ── Determine which months to generate ───────────────────────────────────────
// Usage:
//   dotnet run                          → current month only
//   dotnet run -- 202603                → March 2026 only
//   dotnet run -- 202603 202604 202605  → multiple specific months
List<DateTime> months;

if (args.Length == 0)
{
    months = [DateTime.Today];
}
else
{
    months = [];
    foreach (string arg in args)
    {
        if (arg.Length != 6 || !int.TryParse(arg, out _))
        {
            Console.Error.WriteLine($"Invalid argument '{arg}'. Expected format: YYYYMM (e.g. 202605)");
            return;
        }
        int year  = int.Parse(arg[..4]);
        int month = int.Parse(arg[4..]);
        months.Add(new DateTime(year, month, 1));
    }
}

// ── Generate one invoice per month ───────────────────────────────────────────
foreach (DateTime target in months)
{
    int monthsElapsed = (target.Year - anchorYear) * 12 + (target.Month - anchorMonth);
    int invoiceNo     = anchorInvoiceNo + monthsElapsed;

    string date           = $"{target.Month}/{10}/{target.Year}";
    string dueDate        = $"{target.Month}/{25}/{target.Year}";
    string monthName      = target.ToString("MMMM");
    string outputFileName = $"{outputFilePrefix} {target:yyyyMM}";

    Console.WriteLine($"\nInvoice No: {invoiceNo}  |  Date: {date}  |  Due: {dueDate}  |  Month: {monthName}");

    // Build placeholder map from appsettings + auto-computed values
    Dictionary<string, string> placeholders = config
        .GetSection("Invoice:Placeholders")
        .GetChildren()
        .ToDictionary(
            x => $"{{{{{x.Key}}}}}",
            x => x.Value ?? string.Empty
        );

    placeholders["{{Date}}"]      = date;
    placeholders["{{DueDate}}"]   = dueDate;
    placeholders["{{Month}}"]     = monthName;
    placeholders["{{InvoiceNo}}"] = invoiceNo.ToString();

    string outputDocx = Path.Combine(outputFolder, outputFileName + ".docx");

    // Write via a temp file so an open editor lock on the destination never blocks us
    string tempDocx = Path.Combine(outputFolder, $".tmp_{outputFileName}.docx");
    File.Copy(templatePath, tempDocx, overwrite: true);
    ProcessDocx(tempDocx, placeholders);
    File.Move(tempDocx, outputDocx, overwrite: true);

    Console.WriteLine($"Output:    {outputDocx}");

    string? pdfPath = ConvertToPdf(outputDocx, outputFolder);
    if (pdfPath is not null)
        Console.WriteLine($"PDF:       {pdfPath}");
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/// <summary>
/// Opens the .docx (ZIP) in memory, rewrites word/document.xml with the
/// placeholder values, then saves the modified bytes back to disk.
/// </summary>
static void ProcessDocx(string docxPath, Dictionary<string, string> placeholders)
{
    byte[] zipBytes = File.ReadAllBytes(docxPath);
    byte[] result;

    using (var ms = new MemoryStream())
    {
        ms.Write(zipBytes);

        using (var archive = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            ZipArchiveEntry entry = archive.GetEntry("word/document.xml")
                ?? throw new InvalidOperationException("word/document.xml not found inside the .docx archive.");

            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();

            // Word's spellchecker sometimes splits a placeholder like {{DueDate}}
            // into three separate runs with <w:proofErr> elements in between:
            //   <w:t>{{</w:t></w:r>
            //   <w:proofErr w:type="spellStart"/>
            //   <w:r><w:rPr>…</w:rPr><w:t>DueDate</w:t></w:r>
            //   <w:proofErr w:type="spellEnd"/>
            //   <w:r><w:rPr>…</w:rPr><w:t>}}</w:t></w:r>
            // Collapse those back to a single <w:t>{{DueDate}}</w:t></w:r>.
            xml = NormalizeSplitPlaceholders(xml);

            // Replace every {{Key}} token with its configured value
            foreach (var (token, value) in placeholders)
                xml = xml.Replace(token, SecurityElement.Escape(value));

            // Replace the entry (delete + recreate is the ZipArchive update pattern)
            entry.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry("word/document.xml");
            using var writer = new StreamWriter(newEntry.Open());
            writer.Write(xml);
        }

        result = ms.ToArray();
    }

    File.WriteAllBytes(docxPath, result);
    Console.WriteLine("Placeholders replaced successfully.");
}

/// <summary>
/// Collapses any placeholder that Word has split across multiple runs back into a single run.
///
/// Word produces three distinct split patterns:
///
/// Pattern A — spellcheck split ({{DueDate}}, {{FullName}}, etc.):
///   <w:t>{{</w:t></w:r>
///   <w:proofErr w:type="spellStart"/>
///   <w:r [attrs]><w:rPr>…</w:rPr><w:t>NAME</w:t></w:r>
///   <w:proofErr w:type="spellEnd"/>
///   <w:r [attrs]><w:rPr>…</w:rPr><w:t>}}</w:t></w:r>
///
/// Pattern B — grammar split ({{Amount}}):
///   <w:t>{</w:t></w:r>
///   <w:proofErr w:type="gramEnd"/>
///   <w:r [attrs]><w:rPr>…</w:rPr><w:t>{NAME}}</w:t></w:r>
///
/// Pattern C — mid-name split ({{CustomerAddressLine03}}):
///   <w:t>{{PARTIAL</w:t></w:r>
///   <w:r [attrs]><w:rPr>…</w:rPr><w:t>SUFFIX</w:t></w:r>
///   <w:r [attrs]><w:rPr>…</w:rPr><w:t>}}</w:t></w:r>
/// </summary>
static string NormalizeSplitPlaceholders(string xml)
{
    // Pattern A: spellStart/spellEnd — <w:r> may have optional attributes (w:rsidR etc.)
    const string patternA =
        @"<w:t>\{\{</w:t></w:r>" +
        @"<w:proofErr[^/]*/>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>(\w+)</w:t></w:r>" +
        @"<w:proofErr[^/]*/>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>\}\}</w:t></w:r>";

    xml = Regex.Replace(xml, patternA, "<w:t>{{$1}}</w:t></w:r>", RegexOptions.Singleline);

    // Pattern B: gramEnd — first brace alone, second brace attached to name
    const string patternB =
        @"<w:t>\{</w:t></w:r>" +
        @"<w:proofErr[^/]*/>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>\{(\w+)\}\}</w:t></w:r>";

    xml = Regex.Replace(xml, patternB, "<w:t>{{$1}}</w:t></w:r>", RegexOptions.Singleline);

    // Pattern C: mid-name split — placeholder name broken across consecutive runs with no proofErr
    //   e.g. <w:t>{{CustomerAddressLine0</w:t></w:r>
    //        <w:r ...><w:rPr>...</w:rPr><w:t>3</w:t></w:r>
    //        <w:r ...><w:rPr>...</w:rPr><w:t>}}</w:t></w:r>
    const string patternC =
        @"<w:t>(\{\{\w+)</w:t></w:r>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>(\w+)</w:t></w:r>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>\}\}</w:t></w:r>";

    xml = Regex.Replace(xml, patternC, "<w:t>$1$2}}</w:t></w:r>", RegexOptions.Singleline);

    // Pattern D: mid-name split with closing braces attached to second fragment
    //   e.g. <w:t>{{SwiftCo</w:t></w:r>
    //        <w:proofErr w:type="spellEnd"/>
    //        <w:r ...><w:rPr>...</w:rPr><w:t>de}}</w:t></w:r>
    const string patternD =
        @"<w:t>(\{\{\w+)</w:t></w:r>" +
        @"<w:proofErr[^/]*/>" +
        @"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>(\w+\}\})</w:t></w:r>";

    xml = Regex.Replace(xml, patternD, "<w:t>$1$2</w:t></w:r>", RegexOptions.Singleline);

    return xml;
}

/// <summary>
/// Converts the .docx to PDF using LibreOffice's headless mode.
/// Returns the PDF path on success, or null if LibreOffice is not found.
/// Requires LibreOffice to be installed:
///   brew install --cask libreoffice
/// </summary>
static string? ConvertToPdf(string docxPath, string outputFolder)
{
    string? soffice = FindLibreOffice();

    if (soffice is null)
    {
        Console.WriteLine();
        Console.WriteLine("LibreOffice not found — PDF conversion skipped.");
        Console.WriteLine("Install it with:  brew install --cask libreoffice");
        Console.WriteLine($"The .docx is ready at: {docxPath}");
        return null;
    }

    Console.WriteLine($"Converting to PDF using LibreOffice ({soffice}) …");

    var psi = new ProcessStartInfo
    {
        FileName  = soffice,
        Arguments = $"--headless --convert-to pdf --outdir \"{outputFolder}\" \"{docxPath}\"",
        RedirectStandardOutput = true,
        RedirectStandardError  = true,
        UseShellExecute = false
    };

    using Process process = Process.Start(psi)
        ?? throw new InvalidOperationException("Failed to start LibreOffice process.");

    string stdout = process.StandardOutput.ReadToEnd();
    string stderr = process.StandardError.ReadToEnd();
    process.WaitForExit();

    if (process.ExitCode != 0)
        throw new InvalidOperationException($"LibreOffice exited with code {process.ExitCode}.\n{stderr}");

    if (!string.IsNullOrWhiteSpace(stdout))
        Console.WriteLine(stdout.Trim());

    string pdfName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
    return Path.Combine(outputFolder, pdfName);
}

/// <summary>
/// Locates the LibreOffice <c>soffice</c> binary on macOS (and falls back to
/// common Linux / PATH locations).
/// </summary>
static string? FindLibreOffice()
{
    string[] candidates =
    [
        // macOS default installation path
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        // Homebrew / Linux common paths
        "/usr/local/bin/soffice",
        "/usr/bin/soffice",
        "/opt/libreoffice/program/soffice",
    ];

    foreach (string candidate in candidates)
    {
        if (File.Exists(candidate))
            return candidate;
    }

    // Last resort: check PATH
    var which = Process.Start(new ProcessStartInfo
    {
        FileName  = "which",
        Arguments = "soffice",
        RedirectStandardOutput = true,
        UseShellExecute = false
    });
    which?.WaitForExit();
    string? found = which?.StandardOutput.ReadToEnd().Trim();
    return string.IsNullOrEmpty(found) ? null : found;
}
