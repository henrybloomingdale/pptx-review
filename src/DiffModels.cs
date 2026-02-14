using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PptxReview;

// ── Top-level diff result ──────────────────────────────────────

public class PptxDiffResult
{
    [JsonPropertyName("old_file")]
    public string OldFile { get; set; } = "";

    [JsonPropertyName("new_file")]
    public string NewFile { get; set; } = "";

    [JsonPropertyName("metadata")]
    public PptxMetadataDiff Metadata { get; set; } = new();

    [JsonPropertyName("slides")]
    public SlidesDiff Slides { get; set; } = new();

    [JsonPropertyName("summary")]
    public PptxDiffSummary Summary { get; set; } = new();
}

// ── Metadata diff ──────────────────────────────────────────────

public class PptxMetadataDiff
{
    [JsonPropertyName("changes")]
    public List<PptxFieldChange> Changes { get; set; } = new();
}

public class PptxFieldChange
{
    [JsonPropertyName("field")]
    public string Field { get; set; } = "";

    [JsonPropertyName("old")]
    public object? Old { get; set; }

    [JsonPropertyName("new")]
    public object? New { get; set; }
}

// ── Slides diff ────────────────────────────────────────────────

public class SlidesDiff
{
    [JsonPropertyName("added")]
    public List<SlideEntry> Added { get; set; } = new();

    [JsonPropertyName("deleted")]
    public List<SlideEntry> Deleted { get; set; } = new();

    [JsonPropertyName("modified")]
    public List<SlideModification> Modified { get; set; } = new();
}

public class SlideEntry
{
    [JsonPropertyName("number")]
    public int Number { get; set; }

    [JsonPropertyName("layout")]
    public string Layout { get; set; } = "";

    [JsonPropertyName("text_preview")]
    public string TextPreview { get; set; } = "";
}

public class SlideModification
{
    [JsonPropertyName("old_number")]
    public int OldNumber { get; set; }

    [JsonPropertyName("new_number")]
    public int NewNumber { get; set; }

    [JsonPropertyName("layout")]
    public string Layout { get; set; } = "";

    [JsonPropertyName("shapes_added")]
    public List<ShapeDiffEntry> ShapesAdded { get; set; } = new();

    [JsonPropertyName("shapes_deleted")]
    public List<ShapeDiffEntry> ShapesDeleted { get; set; } = new();

    [JsonPropertyName("shapes_modified")]
    public List<ShapeModification> ShapesModified { get; set; } = new();

    [JsonPropertyName("notes_change")]
    public NotesChange? NotesChange { get; set; }

    [JsonPropertyName("comments_added")]
    public List<string> CommentsAdded { get; set; } = new();

    [JsonPropertyName("comments_deleted")]
    public List<string> CommentsDeleted { get; set; } = new();

    [JsonPropertyName("images_added")]
    public List<string> ImagesAdded { get; set; } = new();

    [JsonPropertyName("images_deleted")]
    public List<string> ImagesDeleted { get; set; } = new();
}

public class ShapeDiffEntry
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}

public class ShapeModification
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("old_text")]
    public string OldText { get; set; } = "";

    [JsonPropertyName("new_text")]
    public string NewText { get; set; } = "";

    [JsonPropertyName("word_changes")]
    public List<PptxWordChange> WordChanges { get; set; } = new();
}

public class PptxWordChange
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";  // "add", "delete", "replace"

    [JsonPropertyName("old")]
    public string? Old { get; set; }

    [JsonPropertyName("new")]
    public string? New { get; set; }

    [JsonPropertyName("position")]
    public int Position { get; set; }
}

public class NotesChange
{
    [JsonPropertyName("old")]
    public string? Old { get; set; }

    [JsonPropertyName("new")]
    public string? New { get; set; }
}

// ── Summary ────────────────────────────────────────────────────

public class PptxDiffSummary
{
    [JsonPropertyName("slides_added")]
    public int SlidesAdded { get; set; }

    [JsonPropertyName("slides_deleted")]
    public int SlidesDeleted { get; set; }

    [JsonPropertyName("slides_modified")]
    public int SlidesModified { get; set; }

    [JsonPropertyName("shapes_added")]
    public int ShapesAdded { get; set; }

    [JsonPropertyName("shapes_deleted")]
    public int ShapesDeleted { get; set; }

    [JsonPropertyName("shapes_modified")]
    public int ShapesModified { get; set; }

    [JsonPropertyName("notes_changed")]
    public int NotesChanged { get; set; }

    [JsonPropertyName("comment_changes")]
    public int CommentChanges { get; set; }

    [JsonPropertyName("image_changes")]
    public int ImageChanges { get; set; }

    [JsonPropertyName("metadata_changes")]
    public int MetadataChanges { get; set; }

    [JsonPropertyName("identical")]
    public bool Identical { get; set; }
}

// ── Extracted presentation data (used internally) ──────────────

public class PresentationExtraction
{
    public string FileName { get; set; } = "";
    public PresentationMetadata Metadata { get; set; } = new();
    public List<ExtractedSlide> Slides { get; set; } = new();
}

public class PresentationMetadata
{
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? LastModifiedBy { get; set; }
    public string? Created { get; set; }
    public string? Modified { get; set; }
    public int SlideCount { get; set; }
}

public class ExtractedSlide
{
    public int Number { get; set; }
    public string Layout { get; set; } = "";
    public List<ExtractedShape> Shapes { get; set; } = new();
    public string? Notes { get; set; }
    public List<string> Comments { get; set; } = new();
    public List<ExtractedImage> Images { get; set; } = new();

    /// <summary>
    /// Returns concatenated text from all shapes for similarity matching.
    /// </summary>
    public string GetAllText()
    {
        var parts = new List<string>();
        foreach (var s in Shapes)
        {
            if (!string.IsNullOrWhiteSpace(s.Text))
                parts.Add(s.Text);
        }
        return string.Join(" ", parts);
    }
}

public class ExtractedShape
{
    public string Name { get; set; } = "";
    public string Type { get; set; } = "";
    public string Text { get; set; } = "";
}

public class ExtractedImage
{
    public string FileName { get; set; } = "";
    public string ContentType { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public long SizeBytes { get; set; }
}
