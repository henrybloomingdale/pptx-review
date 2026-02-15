using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace PptxReview;

/// <summary>
/// JSON source generator context for trim-safe / AOT-compatible serialization.
/// </summary>
[JsonSerializable(typeof(EditManifest))]
[JsonSerializable(typeof(ProcessingResult))]
[JsonSerializable(typeof(ReadResult))]
[JsonSerializable(typeof(PptxDiffResult))]
[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    WriteIndented = true
)]
public partial class PptxReviewJsonContext : JsonSerializerContext { }

/// <summary>
/// Root manifest model deserialized from the JSON input.
/// </summary>
public class EditManifest
{
    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("changes")]
    public List<Change>? Changes { get; set; }

    [JsonPropertyName("comments")]
    public List<CommentDef>? Comments { get; set; }
}

/// <summary>
/// A single change operation on the presentation.
/// </summary>
public class Change
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "replace_text";

    [JsonPropertyName("find")]
    public string? Find { get; set; }

    [JsonPropertyName("replace")]
    public string? Replace { get; set; }

    [JsonPropertyName("slide")]
    public int? Slide { get; set; }

    [JsonPropertyName("shape")]
    public string? Shape { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("position")]
    public int? Position { get; set; }

    [JsonPropertyName("layout")]
    public string? Layout { get; set; }
}

/// <summary>
/// A comment associated with a slide.
/// </summary>
public class CommentDef
{
    [JsonPropertyName("slide")]
    public int Slide { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}

/// <summary>
/// Result of processing a single edit or comment.
/// </summary>
public class EditResult
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}

/// <summary>
/// Overall result summary for JSON output mode.
/// </summary>
public class ProcessingResult
{
    [JsonPropertyName("input")]
    public string Input { get; set; } = "";

    [JsonPropertyName("output")]
    public string? Output { get; set; }

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("changes_attempted")]
    public int ChangesAttempted { get; set; }

    [JsonPropertyName("changes_succeeded")]
    public int ChangesSucceeded { get; set; }

    [JsonPropertyName("comments_attempted")]
    public int CommentsAttempted { get; set; }

    [JsonPropertyName("comments_succeeded")]
    public int CommentsSucceeded { get; set; }

    [JsonPropertyName("results")]
    public List<EditResult> Results { get; set; } = new();

    [JsonPropertyName("success")]
    public bool Success { get; set; }
}

// --- Read mode models ---

/// <summary>
/// Top-level result from reading a presentation.
/// </summary>
public class ReadResult
{
    [JsonPropertyName("slides")]
    public List<SlideInfo> Slides { get; set; } = new();

    [JsonPropertyName("slide_count")]
    public int SlideCount { get; set; }
}

/// <summary>
/// Information about a single slide.
/// </summary>
public class SlideInfo
{
    [JsonPropertyName("number")]
    public int Number { get; set; }

    [JsonPropertyName("layout")]
    public string Layout { get; set; } = "";

    [JsonPropertyName("shapes")]
    public List<ShapeInfo> Shapes { get; set; } = new();

    [JsonPropertyName("notes")]
    public string? Notes { get; set; }

    [JsonPropertyName("comments")]
    public List<string> Comments { get; set; } = new();
}

/// <summary>
/// Information about a single shape on a slide.
/// </summary>
public class ShapeInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}
