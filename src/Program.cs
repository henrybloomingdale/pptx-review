using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using PptxReview;

class Program
{
    static int Main(string[] args)
    {
        // Parse arguments
        string? inputPath = null;
        string? manifestPath = null;
        string? outputPath = null;
        string? author = null;
        bool jsonOutput = false;
        bool dryRun = false;
        bool readMode = false;
        bool diffMode = false;
        bool textConvMode = false;
        bool gitSetup = false;
        bool showHelp = false;
        bool showVersion = false;
        var positionalArgs = new List<string>();

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-v":
                case "--version":
                    showVersion = true;
                    break;
                case "-o":
                case "--output":
                    if (i + 1 < args.Length) outputPath = args[++i];
                    break;
                case "--author":
                    if (i + 1 < args.Length) author = args[++i];
                    break;
                case "--json":
                    jsonOutput = true;
                    break;
                case "--dry-run":
                    dryRun = true;
                    break;
                case "--read":
                    readMode = true;
                    break;
                case "--diff":
                    diffMode = true;
                    break;
                case "--textconv":
                    textConvMode = true;
                    break;
                case "--git-setup":
                    gitSetup = true;
                    break;
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                default:
                    if (!args[i].StartsWith("-"))
                        positionalArgs.Add(args[i]);
                    break;
            }
        }

        // Map positional args
        if (positionalArgs.Count >= 1) inputPath = positionalArgs[0];
        if (positionalArgs.Count >= 2) manifestPath = positionalArgs[1];

        if (showVersion)
        {
            Console.WriteLine($"pptx-review {GetVersion()}");
            return 0;
        }

        // ── Git setup ──────────────────────────────────────────────
        if (gitSetup)
        {
            PrintGitSetup();
            return 0;
        }

        if (showHelp || (inputPath == null && !gitSetup))
        {
            PrintUsage();
            return showHelp ? 0 : 1;
        }

        // ── Diff mode ─────────────────────────────────────────────
        if (diffMode)
        {
            if (manifestPath == null)
            {
                Error("--diff requires two files: pptx-review --diff old.pptx new.pptx");
                return 1;
            }

            if (!File.Exists(inputPath!))
            {
                Error($"Old file not found: {inputPath}");
                return 1;
            }
            if (!File.Exists(manifestPath))
            {
                Error($"New file not found: {manifestPath}");
                return 1;
            }

            try
            {
                var oldPres = PresentationExtractor.Extract(inputPath!);
                var newPres = PresentationExtractor.Extract(manifestPath);
                var diffResult = PresentationDiffer.Diff(oldPres, newPres);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(diffResult, PptxReviewJsonContext.Default.PptxDiffResult));
                }
                else
                {
                    PresentationDiffer.PrintHumanReadable(diffResult);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Error($"Diff failed: {ex.Message}");
                return 1;
            }
        }

        // ── TextConv mode ─────────────────────────────────────────
        if (textConvMode)
        {
            if (!File.Exists(inputPath!))
            {
                Error($"File not found: {inputPath}");
                return 1;
            }

            try
            {
                var extraction = PresentationExtractor.Extract(inputPath!);
                Console.Write(PptxTextConv.Convert(extraction));
                return 0;
            }
            catch (Exception ex)
            {
                Error($"TextConv failed: {ex.Message}");
                return 1;
            }
        }

        // Validate input file
        if (!File.Exists(inputPath))
        {
            Error($"Input file not found: {inputPath}");
            return 1;
        }

        // ── Read mode ─────────────────────────────────────────────
        if (readMode)
        {
            try
            {
                var editor = new PresentationEditor(author ?? "Reader");
                var readResult = editor.Read(inputPath);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(readResult, PptxReviewJsonContext.Default.ReadResult));
                }
                else
                {
                    PrintHumanReadResult(readResult);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Error($"Failed to read presentation: {ex.Message}");
                return 1;
            }
        }

        // ── Edit mode (original behavior) ─────────────────────────
        // Read manifest from file or stdin
        string manifestJson;
        if (manifestPath != null)
        {
            if (!File.Exists(manifestPath))
            {
                Error($"Manifest file not found: {manifestPath}");
                return 1;
            }
            manifestJson = File.ReadAllText(manifestPath);
        }
        else if (!Console.IsInputRedirected)
        {
            Error("No manifest file specified and no stdin input.\nUsage: pptx-review <input.pptx> <edits.json> -o <output.pptx>");
            return 1;
        }
        else
        {
            manifestJson = Console.In.ReadToEnd();
        }

        // Default output path
        if (outputPath == null && !dryRun)
        {
            string dir = Path.GetDirectoryName(inputPath) ?? ".";
            string name = Path.GetFileNameWithoutExtension(inputPath);
            outputPath = Path.Combine(dir, $"{name}_edited.pptx");
        }

        // Deserialize manifest (using source-generated context for trim/AOT safety)
        EditManifest manifest;
        try
        {
            manifest = JsonSerializer.Deserialize(manifestJson, PptxReviewJsonContext.Default.EditManifest)
                ?? throw new Exception("Manifest deserialized to null");
        }
        catch (Exception ex)
        {
            Error($"Failed to parse manifest JSON: {ex.Message}");
            return 1;
        }

        // Resolve author (CLI flag > manifest > default)
        string effectiveAuthor = author ?? manifest.Author ?? "Reviewer";

        // Process
        var presentationEditor = new PresentationEditor(effectiveAuthor);
        ProcessingResult result;

        try
        {
            result = presentationEditor.Process(inputPath, outputPath ?? "", manifest, dryRun);
        }
        catch (Exception ex)
        {
            Error($"Processing failed: {ex.Message}");
            return 1;
        }

        // Output
        if (jsonOutput)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, PptxReviewJsonContext.Default.ProcessingResult));
        }
        else
        {
            PrintHumanResult(result, dryRun);
        }

        return result.Success ? 0 : 1;
    }

    static void PrintUsage()
    {
        Console.Error.WriteLine(@"pptx-review — Read, write, and diff PowerPoint presentations

Usage:
  pptx-review <input.pptx> --read [--json]              Read presentation content
  pptx-review <input.pptx> <edits.json> [options]       Apply edits to presentation
  pptx-review --diff <old.pptx> <new.pptx> [--json]     Semantic presentation diff
  pptx-review --textconv <file.pptx>                     Git textconv (normalized text)
  pptx-review --git-setup                                Print git configuration
  cat edits.json | pptx-review <input.pptx> [options]

Diff & Git Integration:
  --diff                 Compare two presentations semantically (slides, shapes,
                         speaker notes, comments, images, metadata)
  --textconv             Output normalized text for use as git diff textconv driver
  --git-setup            Print .gitattributes and .gitconfig setup instructions

Read/Write Options:
  --read                 Read mode: extract slides, shapes, notes, comments
  -o, --output <path>    Output file path (default: <input>_edited.pptx)
  --author <name>        Author name (overrides manifest author)
  --json                 Output results as JSON
  --dry-run              Validate manifest without modifying
  -v, --version          Show version
  -h, --help             Show this help

Change Types:
  replace_text   Find and replace text (optionally on specific slide)
  set_text       Set text of a named shape on a slide
  set_notes      Set or replace speaker notes for a slide
  delete_slide   Delete a slide by number
  duplicate_slide Duplicate a slide (optionally at position)
  reorder_slide  Move a slide to a new position
  add_slide      Add a new slide (optional layout and position)

JSON Manifest Format:
  {
    ""author"": ""Reviewer Name"",
    ""changes"": [
      { ""type"": ""replace_text"", ""find"": ""old"", ""replace"": ""new"" },
      { ""type"": ""set_text"", ""slide"": 1, ""shape"": ""Title 1"", ""text"": ""New Title"" },
      { ""type"": ""set_notes"", ""slide"": 1, ""text"": ""Speaker notes"" },
      { ""type"": ""delete_slide"", ""slide"": 5 },
      { ""type"": ""add_slide"", ""layout"": ""Blank"", ""position"": 3 }
    ],
    ""comments"": [
      { ""slide"": 1, ""text"": ""This slide needs work"" }
    ]
  }");
    }

    static void PrintGitSetup()
    {
        Console.WriteLine(@"Git Integration for PowerPoint Presentations
═════════════════════════════════════════════

Add to your repository's .gitattributes:

  *.pptx diff=pptx

Add to your .gitconfig (global or per-repo):

  [diff ""pptx""]
      textconv = pptx-review --textconv

Now `git diff` will show meaningful content changes for .pptx files,
including slide text, speaker notes, comments, images, and metadata.

For two-file comparison outside git:

  pptx-review --diff old.pptx new.pptx
  pptx-review --diff old.pptx new.pptx --json
");
    }

    static void PrintHumanResult(ProcessingResult result, bool dryRun)
    {
        string mode = dryRun ? "[DRY RUN] " : "";
        Console.WriteLine($"\n{mode}pptx-review results");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Author:   {result.Author}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        Console.WriteLine();

        foreach (var r in result.Results)
        {
            string icon = r.Success ? "✓" : "✗";
            Console.WriteLine($"  {icon} [{r.Type}] {r.Message}");
        }

        Console.WriteLine();
        if (result.Success)
            Console.WriteLine(dryRun ? "✅ All edits would succeed" : "✅ All edits applied successfully");
        else
            Console.WriteLine("⚠️  Some edits failed (see above)");
    }

    static void PrintHumanReadResult(ReadResult result)
    {
        Console.WriteLine($"\nPresentation: {result.SlideCount} slide(s)");
        Console.WriteLine(new string('─', 50));

        foreach (var slide in result.Slides)
        {
            Console.WriteLine($"\n  Slide {slide.Number} [{slide.Layout}]");
            foreach (var shape in slide.Shapes)
            {
                string text = shape.Text.Length > 80 ? shape.Text.Substring(0, 80) + "…" : shape.Text;
                text = text.Replace("\n", " ");
                Console.WriteLine($"    {shape.Name}: {text}");
            }
            if (slide.Notes != null)
            {
                string notes = slide.Notes.Length > 80 ? slide.Notes.Substring(0, 80) + "…" : slide.Notes;
                Console.WriteLine($"    [Notes] {notes}");
            }
            if (slide.Comments.Count > 0)
            {
                foreach (var comment in slide.Comments)
                    Console.WriteLine($"    [Comment] {comment}");
            }
        }
        Console.WriteLine();
    }

    static string GetVersion()
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        var ver = asm.GetName().Version;
        return ver != null ? $"{ver.Major}.{ver.Minor}.{ver.Build}" : "1.0.0";
    }

    static void Error(string msg) => Console.Error.WriteLine($"Error: {msg}");
}
