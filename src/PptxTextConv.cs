using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PptxReview;

/// <summary>
/// Produces a normalized text representation of a .pptx file suitable for
/// use as a git textconv driver. This allows `git diff` to show meaningful
/// changes for PowerPoint presentations.
///
/// Output format:
/// === METADATA ===
/// Slides: 10
///
/// === SLIDES ===
/// --- Slide 1 [Title Slide] ---
///   Title 1: "Presentation Title"
///   Subtitle 2: "Author Name"
///   [Notes] Speaker notes here
///
/// --- Slide 2 [Two Content] ---
///   Title 1: "Methods"
///   Content Placeholder 2: "Bullet point 1..."
///   [Comment] reviewer: Check this slide
/// </summary>
public static class PptxTextConv
{
    public static string Convert(PresentationExtraction pres)
    {
        var sb = new StringBuilder();

        // ── Metadata ───────────────────────────────────────
        sb.AppendLine("=== METADATA ===");
        if (pres.Metadata.Title != null)
            sb.AppendLine($"Title: {pres.Metadata.Title}");
        if (pres.Metadata.Author != null)
            sb.AppendLine($"Author: {pres.Metadata.Author}");
        if (pres.Metadata.LastModifiedBy != null)
            sb.AppendLine($"LastModifiedBy: {pres.Metadata.LastModifiedBy}");
        if (pres.Metadata.Modified != null)
            sb.AppendLine($"Modified: {pres.Metadata.Modified}");
        sb.AppendLine($"Slides: {pres.Metadata.SlideCount}");
        sb.AppendLine();

        // ── Slides ─────────────────────────────────────────
        sb.AppendLine("=== SLIDES ===");

        foreach (var slide in pres.Slides)
        {
            sb.AppendLine($"--- Slide {slide.Number} [{slide.Layout}] ---");

            // Shapes
            foreach (var shape in slide.Shapes)
            {
                if (string.IsNullOrWhiteSpace(shape.Text) && shape.Type != "Picture")
                    continue;

                if (shape.Type == "Picture")
                {
                    sb.AppendLine($"  [Picture] {shape.Name}");
                }
                else
                {
                    // Normalize multiline text for display
                    string text = shape.Text.Replace("\n", " ¶ ");
                    sb.AppendLine($"  {shape.Name}: \"{text}\"");
                }
            }

            // Speaker notes
            if (slide.Notes != null)
            {
                string notes = slide.Notes.Replace("\n", " ¶ ");
                sb.AppendLine($"  [Notes] {notes}");
            }

            // Comments
            foreach (var comment in slide.Comments)
            {
                sb.AppendLine($"  [Comment] {comment}");
            }

            // Images
            foreach (var img in slide.Images)
            {
                string hash = img.Sha256.Length >= 12 ? img.Sha256[..12] : img.Sha256;
                sb.AppendLine($"  [Image] {img.FileName} ({img.ContentType}, {img.SizeBytes} bytes, sha256:{hash}...)");
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }
}
