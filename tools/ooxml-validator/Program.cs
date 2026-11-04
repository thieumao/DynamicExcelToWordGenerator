// Usage: dotnet run --project tools/ooxml-validator -- [--format Office2019] <file.docx|folder> ...
// Prints one JSON object per validation error and exits with 1 if any file has errors.
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var format = FileFormatVersions.Office2019;
var inputs = new List<string>();
for (var i = 0; i < args.Length; i++)
{
    if (args[i] == "--format" && i + 1 < args.Length)
    {
        format = Enum.Parse<FileFormatVersions>(args[++i]);
    }
    else
    {
        inputs.Add(args[i]);
    }
}

var files = inputs.SelectMany(p => Directory.Exists(p)
        ? Directory.EnumerateFiles(p, "*.docx", SearchOption.AllDirectories)
        : new[] { p })
    .Where(f => !Path.GetFileName(f).StartsWith("~$"))
    .OrderBy(f => f, StringComparer.Ordinal)
    .ToList();

var validator = new OpenXmlValidator(format);
var failed = 0;
foreach (var file in files)
{
    try
    {
        using var doc = WordprocessingDocument.Open(file, false);
        var errors = validator.Validate(doc).Take(200).ToList();
        if (errors.Count > 0)
        {
            failed++;
        }
        foreach (var e in errors)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                file,
                part = e.Part?.Uri.ToString(),
                path = e.Path?.XPath,
                id = e.Id,
                type = e.ErrorType.ToString(),
                description = e.Description,
            }));
        }
    }
    catch (Exception ex)
    {
        failed++;
        Console.WriteLine(JsonSerializer.Serialize(new { file, id = "OpenFailed", description = ex.Message }));
    }
}

Console.Error.WriteLine($"validated {files.Count} file(s) against {format}: {failed} with errors");
return failed > 0 ? 1 : 0;
