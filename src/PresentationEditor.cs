using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PText = DocumentFormat.OpenXml.Presentation.Text;
using PPosition = DocumentFormat.OpenXml.Presentation.Position;

namespace PptxReview;

/// <summary>
/// Core editing engine for programmatic PowerPoint editing using the Open XML SDK.
/// Supports text replacement, shape text setting, speaker notes, slide manipulation,
/// and comments.
/// </summary>
public class PresentationEditor
{
    private readonly string _author;
    private readonly string _dateStr;

    public PresentationEditor(string author, DateTime? date = null)
    {
        _author = author;
        _dateStr = (date ?? DateTime.UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ");
    }

    /// <summary>
    /// Read presentation content and return structured data.
    /// </summary>
    public ReadResult Read(string inputPath)
    {
        using var doc = PresentationDocument.Open(inputPath, false);
        var presentationPart = doc.PresentationPart
            ?? throw new Exception("No PresentationPart found");
        var presentation = presentationPart.Presentation;
        var slideIdList = presentation.SlideIdList
            ?? throw new Exception("No SlideIdList found");

        var result = new ReadResult();
        int slideNumber = 0;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            slideNumber++;
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);

            var slideInfo = new SlideInfo
            {
                Number = slideNumber,
                Layout = GetSlideLayoutName(slidePart),
                Notes = GetSlideNotes(slidePart),
                Comments = GetSlideComments(slidePart, slideNumber, presentationPart)
            };

            // Extract shapes with text
            var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
            if (spTree != null)
            {
                foreach (var shape in spTree.Elements<P.Shape>())
                {
                    var nvSpPr = shape.NonVisualShapeProperties;
                    var name = nvSpPr?.NonVisualDrawingProperties?.Name?.Value ?? "";
                    var textBody = shape.TextBody;

                    if (textBody != null)
                    {
                        string text = GetTextFromTextBody(textBody);
                        slideInfo.Shapes.Add(new ShapeInfo
                        {
                            Name = name,
                            Type = "textbox",
                            Text = text
                        });
                    }
                }
            }

            result.Slides.Add(slideInfo);
        }

        result.SlideCount = slideNumber;
        return result;
    }

    /// <summary>
    /// Process a complete edit manifest against a presentation.
    /// </summary>
    public ProcessingResult Process(string inputPath, string outputPath, EditManifest manifest, bool dryRun = false)
    {
        var result = new ProcessingResult
        {
            Input = inputPath,
            Output = dryRun ? null : outputPath,
            Author = _author,
            ChangesAttempted = manifest.Changes?.Count ?? 0,
            CommentsAttempted = manifest.Comments?.Count ?? 0
        };

        if (!dryRun)
            File.Copy(inputPath, outputPath, true);

        string workPath = dryRun ? CreateTempCopy(inputPath) : outputPath;

        try
        {
            using var doc = PresentationDocument.Open(workPath, true);
            var presentationPart = doc.PresentationPart
                ?? throw new Exception("No PresentationPart found");

            // --- Comments first ---
            if (manifest.Comments != null)
            {
                for (int i = 0; i < manifest.Comments.Count; i++)
                {
                    var cdef = manifest.Comments[i];
                    var er = new EditResult { Index = i, Type = "comment" };

                    if (string.IsNullOrEmpty(cdef.Text))
                    {
                        er.Success = false;
                        er.Message = "Empty comment text";
                    }
                    else if (dryRun)
                    {
                        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
                        bool valid = slideIds != null && cdef.Slide >= 1 && cdef.Slide <= slideIds.Count;
                        er.Success = valid;
                        er.Message = valid
                            ? $"Slide {cdef.Slide} exists"
                            : $"Slide {cdef.Slide} out of range";
                    }
                    else
                    {
                        bool ok = AddComment(presentationPart, cdef.Slide, cdef.Text);
                        er.Success = ok;
                        er.Message = ok
                            ? $"Comment added to slide {cdef.Slide}"
                            : $"Failed to add comment to slide {cdef.Slide}";
                    }

                    result.Results.Add(er);
                    if (er.Success) result.CommentsSucceeded++;
                }
            }

            // --- Changes ---
            if (manifest.Changes != null)
            {
                for (int i = 0; i < manifest.Changes.Count; i++)
                {
                    var change = manifest.Changes[i];
                    var er = new EditResult { Index = i, Type = change.Type };

                    try
                    {
                        switch (change.Type.ToLowerInvariant())
                        {
                            case "replace_text":
                                er = ProcessReplaceText(presentationPart, change, i, dryRun);
                                break;
                            case "set_text":
                                er = ProcessSetText(presentationPart, change, i, dryRun);
                                break;
                            case "set_notes":
                                er = ProcessSetNotes(presentationPart, change, i, dryRun);
                                break;
                            case "delete_slide":
                                er = ProcessDeleteSlide(presentationPart, change, i, dryRun);
                                break;
                            case "duplicate_slide":
                                er = ProcessDuplicateSlide(presentationPart, change, i, dryRun);
                                break;
                            case "reorder_slide":
                                er = ProcessReorderSlide(presentationPart, change, i, dryRun);
                                break;
                            case "add_slide":
                                er = ProcessAddSlide(presentationPart, change, i, dryRun);
                                break;
                            default:
                                er.Success = false;
                                er.Message = $"Unknown change type: {change.Type}";
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        er.Success = false;
                        er.Message = $"Error: {ex.Message}";
                    }

                    result.Results.Add(er);
                    if (er.Success) result.ChangesSucceeded++;
                }
            }

            presentationPart.Presentation.Save();
        }
        finally
        {
            if (dryRun && File.Exists(workPath))
                File.Delete(workPath);
        }

        result.Success = result.ChangesSucceeded == result.ChangesAttempted
                      && result.CommentsSucceeded == result.CommentsAttempted;

        return result;
    }

    #region Change Processors

    private EditResult ProcessReplaceText(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "replace_text" };

        if (string.IsNullOrEmpty(change.Find))
        {
            er.Success = false;
            er.Message = "Missing 'find' field";
            return er;
        }
        if (change.Replace == null)
        {
            er.Success = false;
            er.Message = "Missing 'replace' field";
            return er;
        }

        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        if (slideIds == null || slideIds.Count == 0)
        {
            er.Success = false;
            er.Message = "No slides found";
            return er;
        }

        int totalReplaced = 0;

        if (change.Slide.HasValue)
        {
            // Replace in specific slide only
            int slideIdx = change.Slide.Value - 1;
            if (slideIdx < 0 || slideIdx >= slideIds.Count)
            {
                er.Success = false;
                er.Message = $"Slide {change.Slide.Value} out of range (1-{slideIds.Count})";
                return er;
            }

            var slidePart = (SlidePart)presPart.GetPartById(slideIds[slideIdx].RelationshipId!);
            if (dryRun)
            {
                bool found = FindTextInSlide(slidePart, change.Find);
                er.Success = found;
                er.Message = found
                    ? $"Match found on slide {change.Slide.Value}"
                    : $"No match on slide {change.Slide.Value}: \"{Truncate(change.Find, 60)}\"";
            }
            else
            {
                int n = ReplaceTextInSlide(slidePart, change.Find, change.Replace);
                totalReplaced += n;
                er.Success = n > 0;
                er.Message = n > 0
                    ? $"Replaced {n} occurrence(s) on slide {change.Slide.Value}"
                    : $"No match on slide {change.Slide.Value}: \"{Truncate(change.Find, 60)}\"";
            }
        }
        else
        {
            // Replace across all slides
            if (dryRun)
            {
                bool found = false;
                for (int si = 0; si < slideIds.Count; si++)
                {
                    var slidePart = (SlidePart)presPart.GetPartById(slideIds[si].RelationshipId!);
                    if (FindTextInSlide(slidePart, change.Find))
                    {
                        found = true;
                        break;
                    }
                }
                er.Success = found;
                er.Message = found
                    ? $"Match found for: \"{Truncate(change.Find, 60)}\""
                    : $"No match for: \"{Truncate(change.Find, 60)}\"";
            }
            else
            {
                for (int si = 0; si < slideIds.Count; si++)
                {
                    var slidePart = (SlidePart)presPart.GetPartById(slideIds[si].RelationshipId!);
                    totalReplaced += ReplaceTextInSlide(slidePart, change.Find, change.Replace);
                }
                er.Success = totalReplaced > 0;
                er.Message = totalReplaced > 0
                    ? $"Replaced {totalReplaced} occurrence(s) across all slides"
                    : $"No match for: \"{Truncate(change.Find, 60)}\"";
            }
        }

        return er;
    }

    private EditResult ProcessSetText(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "set_text" };

        if (!change.Slide.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'slide' field";
            return er;
        }
        if (string.IsNullOrEmpty(change.Shape))
        {
            er.Success = false;
            er.Message = "Missing 'shape' field";
            return er;
        }
        if (change.Text == null)
        {
            er.Success = false;
            er.Message = "Missing 'text' field";
            return er;
        }

        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        int slideIdx = change.Slide.Value - 1;
        if (slideIds == null || slideIdx < 0 || slideIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Slide {change.Slide.Value} out of range";
            return er;
        }

        var slidePart = (SlidePart)presPart.GetPartById(slideIds[slideIdx].RelationshipId!);
        var shape = FindShapeByName(slidePart, change.Shape);

        if (shape == null)
        {
            er.Success = false;
            er.Message = $"Shape \"{change.Shape}\" not found on slide {change.Slide.Value}";
            return er;
        }

        if (dryRun)
        {
            er.Success = true;
            er.Message = $"Shape \"{change.Shape}\" found on slide {change.Slide.Value}";
            return er;
        }

        SetShapeText(shape, change.Text);
        er.Success = true;
        er.Message = $"Set text on shape \"{change.Shape}\" (slide {change.Slide.Value})";
        return er;
    }

    private EditResult ProcessSetNotes(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "set_notes" };

        if (!change.Slide.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'slide' field";
            return er;
        }
        if (change.Text == null)
        {
            er.Success = false;
            er.Message = "Missing 'text' field";
            return er;
        }

        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        int slideIdx = change.Slide.Value - 1;
        if (slideIds == null || slideIdx < 0 || slideIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Slide {change.Slide.Value} out of range";
            return er;
        }

        if (dryRun)
        {
            er.Success = true;
            er.Message = $"Slide {change.Slide.Value} exists";
            return er;
        }

        var slidePart = (SlidePart)presPart.GetPartById(slideIds[slideIdx].RelationshipId!);
        SetSlideNotes(slidePart, change.Text);
        er.Success = true;
        er.Message = $"Set notes on slide {change.Slide.Value}";
        return er;
    }

    private EditResult ProcessDeleteSlide(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "delete_slide" };

        if (!change.Slide.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'slide' field";
            return er;
        }

        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        int slideIdx = change.Slide.Value - 1;
        if (slideIds == null || slideIdx < 0 || slideIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Slide {change.Slide.Value} out of range";
            return er;
        }

        if (dryRun)
        {
            er.Success = true;
            er.Message = $"Slide {change.Slide.Value} would be deleted";
            return er;
        }

        var slideId = slideIds[slideIdx];
        var relId = slideId.RelationshipId!.Value!;
        var slidePart = (SlidePart)presPart.GetPartById(relId);

        // Remove from slide id list
        slideId.Remove();

        // Remove the slide part
        presPart.DeletePart(slidePart);

        er.Success = true;
        er.Message = $"Deleted slide {change.Slide.Value}";
        return er;
    }

    private EditResult ProcessDuplicateSlide(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "duplicate_slide" };

        if (!change.Slide.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'slide' field";
            return er;
        }

        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        int slideIdx = change.Slide.Value - 1;
        if (slideIds == null || slideIdx < 0 || slideIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Slide {change.Slide.Value} out of range";
            return er;
        }

        if (dryRun)
        {
            er.Success = true;
            er.Message = $"Slide {change.Slide.Value} would be duplicated";
            return er;
        }

        var sourceSlidePart = (SlidePart)presPart.GetPartById(slideIds[slideIdx].RelationshipId!);

        // Create new slide part
        var newSlidePart = presPart.AddNewPart<SlidePart>();
        using (var sourceStream = sourceSlidePart.GetStream())
        {
            sourceStream.CopyTo(newSlidePart.GetStream(FileMode.Create));
        }

        // Copy the slide layout relationship
        if (sourceSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
        }

        // Copy other relationships (images, etc.)
        foreach (var rel in sourceSlidePart.Parts)
        {
            if (rel.OpenXmlPart is not SlideLayoutPart)
            {
                try { newSlidePart.AddPart(rel.OpenXmlPart); } catch { /* already added */ }
            }
        }

        // Copy external relationships
        foreach (var extRel in sourceSlidePart.ExternalRelationships)
        {
            newSlidePart.AddExternalRelationship(extRel.RelationshipType, extRel.Uri, extRel.Id);
        }

        // Copy hyperlink relationships
        foreach (var hypRel in sourceSlidePart.HyperlinkRelationships)
        {
            newSlidePart.AddHyperlinkRelationship(hypRel.Uri, hypRel.IsExternal, hypRel.Id);
        }

        // Add to slide list
        string newRelId = presPart.GetIdOfPart(newSlidePart);
        uint maxId = slideIds.Max(s => s.Id?.Value ?? 255);

        var newSlideId = new SlideId
        {
            Id = maxId + 1,
            RelationshipId = newRelId
        };

        var slideIdList = presPart.Presentation.SlideIdList!;

        if (change.Position.HasValue)
        {
            int pos = change.Position.Value - 1;
            var currentIds = slideIdList.Elements<SlideId>().ToList();
            if (pos >= 0 && pos < currentIds.Count)
                slideIdList.InsertBefore(newSlideId, currentIds[pos]);
            else
                slideIdList.Append(newSlideId);
        }
        else
        {
            // Insert after the source slide
            slideIdList.InsertAfter(newSlideId, slideIds[slideIdx]);
        }

        er.Success = true;
        er.Message = $"Duplicated slide {change.Slide.Value}";
        return er;
    }

    private EditResult ProcessReorderSlide(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "reorder_slide" };

        if (!change.Slide.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'slide' field";
            return er;
        }
        if (!change.Position.HasValue)
        {
            er.Success = false;
            er.Message = "Missing 'position' field";
            return er;
        }

        var slideIdList = presPart.Presentation.SlideIdList;
        var slideIds = slideIdList?.Elements<SlideId>().ToList();
        int slideIdx = change.Slide.Value - 1;
        int targetIdx = change.Position.Value - 1;

        if (slideIds == null || slideIdx < 0 || slideIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Slide {change.Slide.Value} out of range";
            return er;
        }
        if (targetIdx < 0 || targetIdx >= slideIds.Count)
        {
            er.Success = false;
            er.Message = $"Position {change.Position.Value} out of range";
            return er;
        }

        if (dryRun)
        {
            er.Success = true;
            er.Message = $"Slide {change.Slide.Value} would move to position {change.Position.Value}";
            return er;
        }

        var slideId = slideIds[slideIdx];
        slideId.Remove();

        // Refresh after removal
        var remaining = slideIdList!.Elements<SlideId>().ToList();
        if (targetIdx >= remaining.Count)
        {
            slideIdList.Append(slideId);
        }
        else
        {
            slideIdList.InsertBefore(slideId, remaining[targetIdx]);
        }

        er.Success = true;
        er.Message = $"Moved slide {change.Slide.Value} to position {change.Position.Value}";
        return er;
    }

    private EditResult ProcessAddSlide(PresentationPart presPart, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "add_slide" };

        if (dryRun)
        {
            er.Success = true;
            er.Message = "New slide would be added";
            return er;
        }

        // Find matching layout or use first available
        SlideLayoutPart? layoutPart = FindLayoutPart(presPart, change.Layout);
        if (layoutPart == null)
        {
            er.Success = false;
            er.Message = $"No slide layout found" + (change.Layout != null ? $": \"{change.Layout}\"" : "");
            return er;
        }

        // Create new slide
        var newSlidePart = presPart.AddNewPart<SlidePart>();
        var slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(new A.TransformGroup())
                )
            )
        );
        newSlidePart.Slide = slide;
        newSlidePart.AddPart(layoutPart);
        slide.Save();

        // Add to slide list
        var slideIdList = presPart.Presentation.SlideIdList!;
        var existingIds = slideIdList.Elements<SlideId>().ToList();
        uint maxId = existingIds.Count > 0 ? existingIds.Max(s => s.Id?.Value ?? 255) : 255;

        var newSlideId = new SlideId
        {
            Id = maxId + 1,
            RelationshipId = presPart.GetIdOfPart(newSlidePart)
        };

        if (change.Position.HasValue)
        {
            int pos = change.Position.Value - 1;
            var currentIds = slideIdList.Elements<SlideId>().ToList();
            if (pos >= 0 && pos < currentIds.Count)
                slideIdList.InsertBefore(newSlideId, currentIds[pos]);
            else
                slideIdList.Append(newSlideId);
        }
        else
        {
            slideIdList.Append(newSlideId);
        }

        er.Success = true;
        string layoutName = change.Layout ?? "default";
        er.Message = $"Added new slide with layout \"{layoutName}\"";
        if (change.Position.HasValue)
            er.Message += $" at position {change.Position.Value}";
        return er;
    }

    #endregion

    #region Core Operations

    /// <summary>
    /// Find and replace text in a slide. Returns count of replacements.
    /// </summary>
    private int ReplaceTextInSlide(SlidePart slidePart, string find, string replace)
    {
        int count = 0;
        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return 0;

        foreach (var shape in spTree.Elements<P.Shape>())
        {
            var textBody = shape.TextBody;
            if (textBody == null) continue;

            foreach (var paragraph in textBody.Elements<A.Paragraph>())
            {
                count += ReplaceTextInParagraph(paragraph, find, replace);
            }
        }

        if (count > 0) slidePart.Slide.Save();
        return count;
    }

    /// <summary>
    /// Replace text within a Drawing paragraph, handling multi-run text spans.
    /// </summary>
    private int ReplaceTextInParagraph(A.Paragraph paragraph, string find, string replace)
    {
        var runs = paragraph.Elements<A.Run>().ToList();
        if (runs.Count == 0) return 0;

        // Build concatenated text and run map
        string fullText = string.Join("", runs.Select(r => r.Text?.Text ?? ""));
        int idx = fullText.IndexOf(find, StringComparison.Ordinal);
        if (idx < 0) return 0;

        int count = 0;

        // Simple case: replace all occurrences
        // Rebuild runs after replacement
        while (idx >= 0)
        {
            count++;

            int charPos = 0;
            var runMap = new List<(A.Run run, int start, int end)>();
            foreach (var run in runs)
            {
                string t = run.Text?.Text ?? "";
                runMap.Add((run, charPos, charPos + t.Length));
                charPos += t.Length;
            }

            int matchEnd = idx + find.Length;
            var affected = runMap.Where(r => r.start < matchEnd && r.end > idx).ToList();
            if (affected.Count == 0) break;

            // Get formatting from first affected run
            var firstRun = affected[0].run;
            var runProps = firstRun.RunProperties?.CloneNode(true) as A.RunProperties;

            // Calculate prefix and suffix
            string prefix = "";
            if (idx > affected[0].start)
            {
                string t = affected[0].run.Text?.Text ?? "";
                prefix = t.Substring(0, idx - affected[0].start);
            }

            string suffix = "";
            var lastAffected = affected[affected.Count - 1];
            if (matchEnd < lastAffected.end)
            {
                string t = lastAffected.run.Text?.Text ?? "";
                suffix = t.Substring(matchEnd - lastAffected.start);
            }

            // Build replacement text
            string newText = prefix + replace + suffix;

            // Replace first affected run's text, remove others
            if (firstRun.Text != null)
                firstRun.Text.Text = newText;
            else
                firstRun.Append(new A.Text(newText));

            for (int i = 1; i < affected.Count; i++)
                affected[i].run.Remove();

            // Recalculate for next occurrence, starting after the replacement to avoid infinite loops
            runs = paragraph.Elements<A.Run>().ToList();
            fullText = string.Join("", runs.Select(r => r.Text?.Text ?? ""));
            int searchFrom = idx + replace.Length;
            idx = searchFrom < fullText.Length
                ? fullText.IndexOf(find, searchFrom, StringComparison.Ordinal)
                : -1;
        }

        return count;
    }

    /// <summary>
    /// Check if text exists in a slide.
    /// </summary>
    private bool FindTextInSlide(SlidePart slidePart, string find)
    {
        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return false;

        foreach (var shape in spTree.Elements<P.Shape>())
        {
            var textBody = shape.TextBody;
            if (textBody == null) continue;

            string shapeText = GetTextFromTextBody(textBody);
            if (shapeText.Contains(find, StringComparison.Ordinal))
                return true;
        }
        return false;
    }

    /// <summary>
    /// Find a shape by name on a slide.
    /// </summary>
    private P.Shape? FindShapeByName(SlidePart slidePart, string shapeName)
    {
        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return null;

        return spTree.Elements<P.Shape>()
            .FirstOrDefault(s =>
                s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == shapeName);
    }

    /// <summary>
    /// Set the text content of a shape, preserving the first run's formatting.
    /// </summary>
    private void SetShapeText(P.Shape shape, string text)
    {
        var textBody = shape.TextBody;
        if (textBody == null)
        {
            textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );
            shape.Append(textBody);
        }

        // Preserve run properties from first run if available
        A.RunProperties? existingProps = null;
        var firstPara = textBody.Elements<A.Paragraph>().FirstOrDefault();
        if (firstPara != null)
        {
            var firstRun = firstPara.Elements<A.Run>().FirstOrDefault();
            existingProps = firstRun?.RunProperties?.CloneNode(true) as A.RunProperties;
        }

        // Remove existing paragraphs
        textBody.RemoveAllChildren<A.Paragraph>();

        // Split text by newlines and create paragraphs
        var lines = text.Split('\n');
        foreach (var line in lines)
        {
            var para = new A.Paragraph();
            var run = new A.Run();
            if (existingProps != null)
                run.RunProperties = existingProps.CloneNode(true) as A.RunProperties;
            run.Text = new A.Text(line);
            para.Append(run);
            textBody.Append(para);
        }
    }

    /// <summary>
    /// Set or replace speaker notes for a slide.
    /// </summary>
    private void SetSlideNotes(SlidePart slidePart, string notesText)
    {
        NotesSlidePart notesPart;

        if (slidePart.NotesSlidePart != null)
        {
            notesPart = slidePart.NotesSlidePart;
        }
        else
        {
            notesPart = slidePart.AddNewPart<NotesSlidePart>();
            notesPart.NotesSlide = new NotesSlide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties()
                    )
                )
            );
        }

        // Find or create the notes text shape
        var spTree = notesPart.NotesSlide.CommonSlideData?.ShapeTree;
        P.Shape? notesShape = null;

        if (spTree != null)
        {
            // Look for the notes placeholder (type = body or idx = 1)
            notesShape = spTree.Elements<P.Shape>().FirstOrDefault(s =>
            {
                var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>();
                return ph != null && (ph.Type?.Value == PlaceholderValues.Body ||
                                     ph.Index?.Value == 1);
            });
        }

        if (notesShape == null)
        {
            // Create a notes text shape
            notesShape = new P.Shape();
            notesShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 2, Name = "Notes Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 }
                )
            );
            notesShape.ShapeProperties = new P.ShapeProperties();
            spTree?.Append(notesShape);
        }

        // Set the text
        var textBody = notesShape.TextBody;
        if (textBody == null)
        {
            textBody = new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle()
            );
            notesShape.Append(textBody);
        }

        textBody.RemoveAllChildren<A.Paragraph>();

        var lines = notesText.Split('\n');
        foreach (var line in lines)
        {
            var para = new A.Paragraph();
            var run = new A.Run();
            run.Text = new A.Text(line);
            para.Append(run);
            textBody.Append(para);
        }

        notesPart.NotesSlide.Save();
    }

    /// <summary>
    /// Add a comment to a specific slide using modern PowerPoint comment format.
    /// </summary>
    private bool AddComment(PresentationPart presPart, int slideNumber, string commentText)
    {
        var slideIds = presPart.Presentation.SlideIdList?.Elements<SlideId>().ToList();
        if (slideIds == null || slideNumber < 1 || slideNumber > slideIds.Count)
            return false;

        // Ensure comment authors part exists
        var authorsPart = presPart.CommentAuthorsPart;
        if (authorsPart == null)
        {
            authorsPart = presPart.AddNewPart<CommentAuthorsPart>();
            authorsPart.CommentAuthorList = new CommentAuthorList();
        }

        // Find or add author
        var authorList = authorsPart.CommentAuthorList;
        var author = authorList.Elements<CommentAuthor>()
            .FirstOrDefault(a => a.Name?.Value == _author);

        uint authorId;
        uint authorLastIdx;

        if (author != null)
        {
            authorId = author.Id?.Value ?? 0;
            authorLastIdx = (author.LastIndex?.Value ?? 0) + 1;
            author.LastIndex = authorLastIdx;
        }
        else
        {
            var existingAuthors = authorList.Elements<CommentAuthor>().ToList();
            authorId = existingAuthors.Count > 0
                ? existingAuthors.Max(a => a.Id?.Value ?? 0) + 1
                : 0;
            authorLastIdx = 1;

            authorList.Append(new CommentAuthor
            {
                Id = authorId,
                Name = _author,
                Initials = _author.Length > 0 ? _author[0].ToString() : "R",
                LastIndex = authorLastIdx,
                ColorIndex = 0
            });
        }
        authorsPart.CommentAuthorList.Save();

        // Get the slide part
        var slidePart = (SlidePart)presPart.GetPartById(slideIds[slideNumber - 1].RelationshipId!);

        // Ensure slide comments part exists
        var commentsPart = slidePart.SlideCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = slidePart.AddNewPart<SlideCommentsPart>();
            commentsPart.CommentList = new CommentList();
        }

        // Add the comment
        var existingComments = commentsPart.CommentList.Elements<Comment>().ToList();
        uint commentIdx = existingComments.Count > 0
            ? existingComments.Max(c => c.Index?.Value ?? 0) + 1
            : 1;

        var comment = new Comment
        {
            AuthorId = authorId,
            DateTime = new DateTimeValue(DateTime.Parse(_dateStr)),
            Index = commentIdx
        };

        // Position at top-left of slide
        comment.Position = new PPosition { X = 0, Y = 0 };
        comment.Text = new PText(commentText);

        commentsPart.CommentList.Append(comment);
        commentsPart.CommentList.Save();

        return true;
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Extract text from a TextBody element (works with both P.TextBody and A.TextBody).
    /// </summary>
    private string GetTextFromTextBody(OpenXmlCompositeElement textBody)
    {
        var paragraphs = textBody.Elements<A.Paragraph>();
        return string.Join("\n", paragraphs.Select(p =>
            string.Join("", p.Elements<A.Run>().Select(r => r.Text?.Text ?? ""))
        ));
    }

    /// <summary>
    /// Get slide layout name.
    /// </summary>
    private string GetSlideLayoutName(SlidePart slidePart)
    {
        try
        {
            return slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    /// <summary>
    /// Get speaker notes text from a slide.
    /// </summary>
    private string? GetSlideNotes(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart == null) return null;

        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return null;

        foreach (var shape in spTree.Elements<P.Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>();
            if (ph != null && (ph.Type?.Value == PlaceholderValues.Body || ph.Index?.Value == 1))
            {
                var textBody = shape.TextBody;
                if (textBody != null)
                {
                    string text = GetTextFromTextBody(textBody);
                    return string.IsNullOrWhiteSpace(text) ? null : text;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Get comments for a specific slide.
    /// </summary>
    private List<string> GetSlideComments(SlidePart slidePart, int slideNumber, PresentationPart presPart)
    {
        var comments = new List<string>();
        var commentsPart = slidePart.SlideCommentsPart;
        if (commentsPart?.CommentList == null) return comments;

        // Build author lookup
        var authorLookup = new Dictionary<uint, string>();
        var authorsPart = presPart.CommentAuthorsPart;
        if (authorsPart?.CommentAuthorList != null)
        {
            foreach (var author in authorsPart.CommentAuthorList.Elements<CommentAuthor>())
            {
                if (author.Id?.Value != null)
                    authorLookup[author.Id.Value] = author.Name?.Value ?? "Unknown";
            }
        }

        foreach (var comment in commentsPart.CommentList.Elements<Comment>())
        {
            string authorName = comment.AuthorId?.Value != null
                && authorLookup.TryGetValue(comment.AuthorId.Value, out var name)
                    ? name : "Unknown";
            string text = comment.GetFirstChild<PText>()?.Text ?? "";
            comments.Add($"[{authorName}] {text}");
        }

        return comments;
    }

    /// <summary>
    /// Find a layout part by name, or return the first available layout.
    /// </summary>
    private SlideLayoutPart? FindLayoutPart(PresentationPart presPart, string? layoutName)
    {
        foreach (var masterPart in presPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                if (layoutName == null)
                    return layoutPart; // Return first available

                string name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? "";
                if (name.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    return layoutPart;
            }
        }

        // If specific name not found, return first available
        if (layoutName != null)
        {
            foreach (var masterPart in presPart.SlideMasterParts)
            {
                var first = masterPart.SlideLayoutParts.FirstOrDefault();
                if (first != null) return first;
            }
        }

        return null;
    }

    private static string CreateTempCopy(string path)
    {
        string tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"pptx-review-{Guid.NewGuid()}.pptx");
        File.Copy(path, tmp);
        return tmp;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "…";

    #endregion
}
