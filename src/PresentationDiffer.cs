using System;
using System.Collections.Generic;
using System.Linq;

namespace PptxReview;

/// <summary>
/// Compares two PresentationExtractions and produces a semantic PptxDiffResult
/// covering slides, shapes, speaker notes, comments, images, and metadata.
/// </summary>
public static class PresentationDiffer
{
    public static PptxDiffResult Diff(PresentationExtraction oldPres, PresentationExtraction newPres)
    {
        var result = new PptxDiffResult
        {
            OldFile = oldPres.FileName,
            NewFile = newPres.FileName
        };

        result.Metadata = DiffMetadata(oldPres.Metadata, newPres.Metadata);
        result.Slides = DiffSlides(oldPres.Slides, newPres.Slides);

        // Build summary
        int notesChanged = result.Slides.Modified.Count(m => m.NotesChange != null);
        int commentChanges = result.Slides.Modified.Sum(m => m.CommentsAdded.Count + m.CommentsDeleted.Count);
        int imageChanges = result.Slides.Modified.Sum(m => m.ImagesAdded.Count + m.ImagesDeleted.Count);

        result.Summary = new PptxDiffSummary
        {
            SlidesAdded = result.Slides.Added.Count,
            SlidesDeleted = result.Slides.Deleted.Count,
            SlidesModified = result.Slides.Modified.Count,
            ShapesAdded = result.Slides.Modified.Sum(m => m.ShapesAdded.Count),
            ShapesDeleted = result.Slides.Modified.Sum(m => m.ShapesDeleted.Count),
            ShapesModified = result.Slides.Modified.Sum(m => m.ShapesModified.Count),
            NotesChanged = notesChanged,
            CommentChanges = commentChanges,
            ImageChanges = imageChanges,
            MetadataChanges = result.Metadata.Changes.Count,
            Identical = result.Metadata.Changes.Count == 0
                     && result.Slides.Added.Count == 0
                     && result.Slides.Deleted.Count == 0
                     && result.Slides.Modified.Count == 0
        };

        return result;
    }

    // ── Metadata ───────────────────────────────────────────────

    private static PptxMetadataDiff DiffMetadata(PresentationMetadata oldMeta, PresentationMetadata newMeta)
    {
        var diff = new PptxMetadataDiff();

        CompareField(diff.Changes, "title", oldMeta.Title, newMeta.Title);
        CompareField(diff.Changes, "author", oldMeta.Author, newMeta.Author);
        CompareField(diff.Changes, "last_modified_by", oldMeta.LastModifiedBy, newMeta.LastModifiedBy);
        CompareField(diff.Changes, "created", oldMeta.Created, newMeta.Created);
        CompareField(diff.Changes, "modified", oldMeta.Modified, newMeta.Modified);

        if (oldMeta.SlideCount != newMeta.SlideCount)
            diff.Changes.Add(new PptxFieldChange { Field = "slide_count", Old = oldMeta.SlideCount, New = newMeta.SlideCount });

        return diff;
    }

    private static void CompareField(List<PptxFieldChange> changes, string name, string? oldVal, string? newVal)
    {
        if (oldVal != newVal)
            changes.Add(new PptxFieldChange { Field = name, Old = oldVal, New = newVal });
    }

    // ── Slides ─────────────────────────────────────────────────

    private static SlidesDiff DiffSlides(List<ExtractedSlide> oldSlides, List<ExtractedSlide> newSlides)
    {
        var diff = new SlidesDiff();

        // Align slides using content similarity
        var alignment = AlignSlides(oldSlides, newSlides);

        foreach (var (oi, ni) in alignment)
        {
            if (oi < 0)
            {
                // Added slide
                var ns = newSlides[ni];
                diff.Added.Add(new SlideEntry
                {
                    Number = ns.Number,
                    Layout = ns.Layout,
                    TextPreview = Trunc(ns.GetAllText(), 120)
                });
            }
            else if (ni < 0)
            {
                // Deleted slide
                var os = oldSlides[oi];
                diff.Deleted.Add(new SlideEntry
                {
                    Number = os.Number,
                    Layout = os.Layout,
                    TextPreview = Trunc(os.GetAllText(), 120)
                });
            }
            else
            {
                // Matched — check for modifications
                var oldSlide = oldSlides[oi];
                var newSlide = newSlides[ni];
                var mod = DiffSingleSlide(oldSlide, newSlide);

                if (mod != null)
                    diff.Modified.Add(mod);
            }
        }

        return diff;
    }

    /// <summary>
    /// Align slides using LCS on content similarity.
    /// Returns list of (oldIndex, newIndex) pairs. -1 means added/deleted.
    /// </summary>
    private static List<(int oldIdx, int newIdx)> AlignSlides(
        List<ExtractedSlide> oldSlides, List<ExtractedSlide> newSlides)
    {
        int m = oldSlides.Count;
        int n = newSlides.Count;

        // LCS table using content similarity
        var dp = new int[m + 1, n + 1];
        for (int i = 1; i <= m; i++)
        {
            for (int j = 1; j <= n; j++)
            {
                if (AreSimilar(oldSlides[i - 1], newSlides[j - 1]))
                    dp[i, j] = dp[i - 1, j - 1] + 1;
                else
                    dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
            }
        }

        // Backtrack to get matched pairs
        var matched = new List<(int, int)>();
        int oi = m, ni = n;

        while (oi > 0 && ni > 0)
        {
            if (AreSimilar(oldSlides[oi - 1], newSlides[ni - 1]))
            {
                matched.Add((oi - 1, ni - 1));
                oi--; ni--;
            }
            else if (dp[oi - 1, ni] > dp[oi, ni - 1])
            {
                oi--;
            }
            else
            {
                ni--;
            }
        }

        matched.Reverse();

        // Build full alignment including unmatched
        var result = new List<(int, int)>();
        int mi = 0, oPtr = 0, nPtr = 0;

        while (mi < matched.Count || oPtr < m || nPtr < n)
        {
            if (mi < matched.Count)
            {
                var (mo, mn) = matched[mi];
                while (oPtr < mo)
                    result.Add((oPtr++, -1));
                while (nPtr < mn)
                    result.Add((-1, nPtr++));
                result.Add((mo, mn));
                oPtr = mo + 1;
                nPtr = mn + 1;
                mi++;
            }
            else
            {
                while (oPtr < m)
                    result.Add((oPtr++, -1));
                while (nPtr < n)
                    result.Add((-1, nPtr++));
                break;
            }
        }

        return result;
    }

    /// <summary>
    /// Two slides are "similar" if they share enough textual content.
    /// Uses Jaccard similarity on words ≥ 0.3 threshold, or same layout + position.
    /// </summary>
    private static bool AreSimilar(ExtractedSlide a, ExtractedSlide b)
    {
        string textA = a.GetAllText();
        string textB = b.GetAllText();

        if (textA == textB && a.Layout == b.Layout) return true;

        // Both empty text but same layout
        if (string.IsNullOrWhiteSpace(textA) && string.IsNullOrWhiteSpace(textB))
            return a.Layout == b.Layout;

        if (string.IsNullOrWhiteSpace(textA) || string.IsNullOrWhiteSpace(textB))
            return false;

        // Jaccard similarity on words
        var wordsA = new HashSet<string>(textA.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        var wordsB = new HashSet<string>(textB.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        if (wordsA.Count == 0 && wordsB.Count == 0) return true;

        int intersection = wordsA.Intersect(wordsB).Count();
        int union = wordsA.Union(wordsB).Count();

        return union > 0 && (double)intersection / union >= 0.3;
    }

    /// <summary>
    /// Compare two matched slides for shape-level, notes, comment, and image differences.
    /// Returns null if slides are identical.
    /// </summary>
    private static SlideModification? DiffSingleSlide(ExtractedSlide oldSlide, ExtractedSlide newSlide)
    {
        var mod = new SlideModification
        {
            OldNumber = oldSlide.Number,
            NewNumber = newSlide.Number,
            Layout = newSlide.Layout
        };

        bool hasChanges = false;

        // ── Shapes ─────────────────────────────────────────
        DiffShapes(oldSlide.Shapes, newSlide.Shapes, mod);
        if (mod.ShapesAdded.Count > 0 || mod.ShapesDeleted.Count > 0 || mod.ShapesModified.Count > 0)
            hasChanges = true;

        // ── Notes ──────────────────────────────────────────
        if (oldSlide.Notes != newSlide.Notes)
        {
            mod.NotesChange = new NotesChange
            {
                Old = oldSlide.Notes,
                New = newSlide.Notes
            };
            hasChanges = true;
        }

        // ── Comments ───────────────────────────────────────
        var oldCommentSet = new HashSet<string>(oldSlide.Comments);
        var newCommentSet = new HashSet<string>(newSlide.Comments);

        mod.CommentsAdded = newSlide.Comments.Where(c => !oldCommentSet.Contains(c)).ToList();
        mod.CommentsDeleted = oldSlide.Comments.Where(c => !newCommentSet.Contains(c)).ToList();
        if (mod.CommentsAdded.Count > 0 || mod.CommentsDeleted.Count > 0)
            hasChanges = true;

        // ── Images ─────────────────────────────────────────
        var oldImageHashes = new HashSet<string>(oldSlide.Images.Select(i => i.Sha256));
        var newImageHashes = new HashSet<string>(newSlide.Images.Select(i => i.Sha256));

        mod.ImagesAdded = newSlide.Images
            .Where(i => !oldImageHashes.Contains(i.Sha256))
            .Select(i => $"{i.FileName} ({i.ContentType}, sha256:{i.Sha256[..Math.Min(12, i.Sha256.Length)]}...)")
            .ToList();
        mod.ImagesDeleted = oldSlide.Images
            .Where(i => !newImageHashes.Contains(i.Sha256))
            .Select(i => $"{i.FileName} ({i.ContentType}, sha256:{i.Sha256[..Math.Min(12, i.Sha256.Length)]}...)")
            .ToList();
        if (mod.ImagesAdded.Count > 0 || mod.ImagesDeleted.Count > 0)
            hasChanges = true;

        return hasChanges ? mod : null;
    }

    /// <summary>
    /// Compare shapes between two slides by matching on shape name.
    /// </summary>
    private static void DiffShapes(List<ExtractedShape> oldShapes, List<ExtractedShape> newShapes, SlideModification mod)
    {
        var oldByName = new Dictionary<string, ExtractedShape>();
        foreach (var s in oldShapes)
        {
            if (!string.IsNullOrEmpty(s.Name) && !oldByName.ContainsKey(s.Name))
                oldByName[s.Name] = s;
        }

        var newByName = new Dictionary<string, ExtractedShape>();
        foreach (var s in newShapes)
        {
            if (!string.IsNullOrEmpty(s.Name) && !newByName.ContainsKey(s.Name))
                newByName[s.Name] = s;
        }

        // Deleted shapes
        foreach (var key in oldByName.Keys.Except(newByName.Keys))
        {
            var s = oldByName[key];
            mod.ShapesDeleted.Add(new ShapeDiffEntry { Name = s.Name, Type = s.Type, Text = s.Text });
        }

        // Added shapes
        foreach (var key in newByName.Keys.Except(oldByName.Keys))
        {
            var s = newByName[key];
            mod.ShapesAdded.Add(new ShapeDiffEntry { Name = s.Name, Type = s.Type, Text = s.Text });
        }

        // Modified shapes (same name, different text)
        foreach (var key in oldByName.Keys.Intersect(newByName.Keys))
        {
            var oldShape = oldByName[key];
            var newShape = newByName[key];

            if (oldShape.Text != newShape.Text)
            {
                mod.ShapesModified.Add(new ShapeModification
                {
                    Name = oldShape.Name,
                    Type = newShape.Type,
                    OldText = oldShape.Text,
                    NewText = newShape.Text,
                    WordChanges = ComputeWordDiff(oldShape.Text, newShape.Text)
                });
            }
        }
    }

    // ── Word-level diff ────────────────────────────────────────

    private static List<PptxWordChange> ComputeWordDiff(string oldText, string newText)
    {
        var oldWords = oldText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var newWords = newText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

        var changes = new List<PptxWordChange>();
        var lcs = WordLCS(oldWords, newWords);

        int oi = 0, ni = 0, li = 0;

        while (oi < oldWords.Length || ni < newWords.Length)
        {
            if (li < lcs.Count)
            {
                var (lo, ln) = lcs[li];

                while (oi < lo)
                {
                    changes.Add(new PptxWordChange { Type = "delete", Old = oldWords[oi], Position = oi });
                    oi++;
                }
                while (ni < ln)
                {
                    changes.Add(new PptxWordChange { Type = "add", New = newWords[ni], Position = ni });
                    ni++;
                }

                oi = lo + 1;
                ni = ln + 1;
                li++;
            }
            else
            {
                while (oi < oldWords.Length)
                {
                    changes.Add(new PptxWordChange { Type = "delete", Old = oldWords[oi], Position = oi });
                    oi++;
                }
                while (ni < newWords.Length)
                {
                    changes.Add(new PptxWordChange { Type = "add", New = newWords[ni], Position = ni });
                    ni++;
                }
            }
        }

        return CollapseToReplace(changes);
    }

    private static List<(int, int)> WordLCS(string[] a, string[] b)
    {
        int m = a.Length, n = b.Length;
        var dp = new int[m + 1, n + 1];

        for (int i = 1; i <= m; i++)
            for (int j = 1; j <= n; j++)
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);

        var result = new List<(int, int)>();
        int oi2 = m, ni2 = n;
        while (oi2 > 0 && ni2 > 0)
        {
            if (a[oi2 - 1] == b[ni2 - 1])
            {
                result.Add((oi2 - 1, ni2 - 1));
                oi2--; ni2--;
            }
            else if (dp[oi2 - 1, ni2] > dp[oi2, ni2 - 1])
                oi2--;
            else
                ni2--;
        }

        result.Reverse();
        return result;
    }

    private static List<PptxWordChange> CollapseToReplace(List<PptxWordChange> changes)
    {
        var result = new List<PptxWordChange>();

        for (int i = 0; i < changes.Count; i++)
        {
            if (i + 1 < changes.Count
                && changes[i].Type == "delete"
                && changes[i + 1].Type == "add")
            {
                result.Add(new PptxWordChange
                {
                    Type = "replace",
                    Old = changes[i].Old,
                    New = changes[i + 1].New,
                    Position = changes[i].Position
                });
                i++; // skip the add
            }
            else
            {
                result.Add(changes[i]);
            }
        }

        return result;
    }

    // ── Human-readable output ──────────────────────────────────

    public static void PrintHumanReadable(PptxDiffResult result)
    {
        Console.WriteLine();
        Console.WriteLine($"pptx-review diff: {result.OldFile} → {result.NewFile}");
        Console.WriteLine(new string('═', 60));

        if (result.Summary.Identical)
        {
            Console.WriteLine("\n  Presentations are identical.");
            return;
        }

        // Metadata
        if (result.Metadata.Changes.Count > 0)
        {
            Console.WriteLine("\nMetadata");
            Console.WriteLine(new string('─', 40));
            foreach (var c in result.Metadata.Changes)
            {
                string oldVal = c.Old?.ToString() ?? "(none)";
                string newVal = c.New?.ToString() ?? "(none)";
                Console.WriteLine($"  {c.Field}: {Trunc(oldVal, 40)} → {Trunc(newVal, 40)}");
            }
        }

        // Added slides
        if (result.Slides.Added.Count > 0)
        {
            Console.WriteLine($"\nSlides Added ({result.Slides.Added.Count})");
            Console.WriteLine(new string('─', 40));
            foreach (var s in result.Slides.Added)
                Console.WriteLine($"  + Slide {s.Number} [{s.Layout}]: {Trunc(s.TextPreview, 72)}");
        }

        // Deleted slides
        if (result.Slides.Deleted.Count > 0)
        {
            Console.WriteLine($"\nSlides Deleted ({result.Slides.Deleted.Count})");
            Console.WriteLine(new string('─', 40));
            foreach (var s in result.Slides.Deleted)
                Console.WriteLine($"  - Slide {s.Number} [{s.Layout}]: {Trunc(s.TextPreview, 72)}");
        }

        // Modified slides
        if (result.Slides.Modified.Count > 0)
        {
            Console.WriteLine($"\nSlides Modified ({result.Slides.Modified.Count})");
            Console.WriteLine(new string('─', 40));

            foreach (var m in result.Slides.Modified)
            {
                string slideLabel = m.OldNumber == m.NewNumber
                    ? $"Slide {m.OldNumber}"
                    : $"Slide {m.OldNumber}→{m.NewNumber}";
                Console.WriteLine($"\n  {slideLabel} [{m.Layout}]:");

                // Shapes added
                foreach (var s in m.ShapesAdded)
                    Console.WriteLine($"    + Shape \"{s.Name}\" ({s.Type}): \"{Trunc(s.Text, 60)}\"");

                // Shapes deleted
                foreach (var s in m.ShapesDeleted)
                    Console.WriteLine($"    - Shape \"{s.Name}\" ({s.Type}): \"{Trunc(s.Text, 60)}\"");

                // Shapes modified
                foreach (var s in m.ShapesModified)
                {
                    Console.WriteLine($"    ~ Shape \"{s.Name}\" ({s.Type}):");
                    Console.WriteLine($"      - \"{Trunc(s.OldText, 60)}\"");
                    Console.WriteLine($"      + \"{Trunc(s.NewText, 60)}\"");

                    if (s.WordChanges.Count > 0 && s.WordChanges.Count <= 5)
                    {
                        foreach (var wc in s.WordChanges)
                        {
                            string desc = wc.Type switch
                            {
                                "replace" => $"\"{wc.Old}\" → \"{wc.New}\"",
                                "add" => $"+ \"{wc.New}\"",
                                "delete" => $"- \"{wc.Old}\"",
                                _ => wc.Type
                            };
                            Console.WriteLine($"        {desc}");
                        }
                    }
                }

                // Notes
                if (m.NotesChange != null)
                {
                    Console.WriteLine($"    Notes:");
                    Console.WriteLine($"      - \"{Trunc(m.NotesChange.Old ?? "(none)", 60)}\"");
                    Console.WriteLine($"      + \"{Trunc(m.NotesChange.New ?? "(none)", 60)}\"");
                }

                // Comments
                foreach (var c in m.CommentsAdded)
                    Console.WriteLine($"    + Comment: {c}");
                foreach (var c in m.CommentsDeleted)
                    Console.WriteLine($"    - Comment: {c}");

                // Images
                foreach (var img in m.ImagesAdded)
                    Console.WriteLine($"    + Image: {img}");
                foreach (var img in m.ImagesDeleted)
                    Console.WriteLine($"    - Image: {img}");
            }
        }

        // Summary
        Console.WriteLine($"\nSummary: {result.Summary.SlidesAdded} added, "
            + $"{result.Summary.SlidesDeleted} deleted, "
            + $"{result.Summary.SlidesModified} modified slides | "
            + $"{result.Summary.ShapesModified} shape text changes, "
            + $"{result.Summary.NotesChanged} notes changes, "
            + $"{result.Summary.CommentChanges} comment changes, "
            + $"{result.Summary.ImageChanges} image changes, "
            + $"{result.Summary.MetadataChanges} metadata changes");
        Console.WriteLine();
    }

    private static string Trunc(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";
}
