using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PText = DocumentFormat.OpenXml.Presentation.Text;

namespace PptxReview;

/// <summary>
/// Extracts all content from a .pptx file into a PresentationExtraction
/// for use by PresentationDiffer and PptxTextConv.
/// </summary>
public static class PresentationExtractor
{
    public static PresentationExtraction Extract(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}");

        var result = new PresentationExtraction
        {
            FileName = Path.GetFileName(path)
        };

        using var doc = PresentationDocument.Open(path, false);
        var presPart = doc.PresentationPart
            ?? throw new Exception("No PresentationPart found");
        var presentation = presPart.Presentation;
        var slideIdList = presentation.SlideIdList
            ?? throw new Exception("No SlideIdList found");

        // Metadata
        result.Metadata = ExtractMetadata(doc, slideIdList);

        // Slides
        int slideNumber = 0;
        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            slideNumber++;
            var slidePart = (SlidePart)presPart.GetPartById(slideId.RelationshipId!);
            result.Slides.Add(ExtractSlide(slidePart, slideNumber, presPart));
        }

        return result;
    }

    private static PresentationMetadata ExtractMetadata(PresentationDocument doc, SlideIdList slideIdList)
    {
        var meta = new PresentationMetadata();
        var props = doc.PackageProperties;
        if (props != null)
        {
            meta.Title = props.Title;
            meta.Author = props.Creator;
            meta.LastModifiedBy = props.LastModifiedBy;
            if (props.Created.HasValue)
                meta.Created = props.Created.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
            if (props.Modified.HasValue)
                meta.Modified = props.Modified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
        }
        meta.SlideCount = slideIdList.Elements<SlideId>().Count();
        return meta;
    }

    private static ExtractedSlide ExtractSlide(SlidePart slidePart, int number, PresentationPart presPart)
    {
        var slide = new ExtractedSlide
        {
            Number = number,
            Layout = GetSlideLayoutName(slidePart),
            Notes = GetSlideNotes(slidePart),
            Comments = GetSlideComments(slidePart, presPart)
        };

        // Extract shapes
        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree != null)
        {
            foreach (var shape in spTree.Elements<P.Shape>())
            {
                var nvSpPr = shape.NonVisualShapeProperties;
                var name = nvSpPr?.NonVisualDrawingProperties?.Name?.Value ?? "";
                var textBody = shape.TextBody;
                string shapeType = GetShapeType(shape);

                string text = "";
                if (textBody != null)
                    text = GetTextFromTextBody(textBody);

                slide.Shapes.Add(new ExtractedShape
                {
                    Name = name,
                    Type = shapeType,
                    Text = text
                });
            }

            // Pictures
            foreach (var pic in spTree.Elements<P.Picture>())
            {
                var nvPicPr = pic.NonVisualPictureProperties;
                var name = nvPicPr?.NonVisualDrawingProperties?.Name?.Value ?? "";
                slide.Shapes.Add(new ExtractedShape
                {
                    Name = name,
                    Type = "Picture",
                    Text = ""
                });
            }

            // Group shapes (flatten)
            foreach (var grp in spTree.Elements<P.GroupShape>())
            {
                foreach (var shape in grp.Elements<P.Shape>())
                {
                    var nvSpPr = shape.NonVisualShapeProperties;
                    var name = nvSpPr?.NonVisualDrawingProperties?.Name?.Value ?? "";
                    var textBody = shape.TextBody;
                    string text = textBody != null ? GetTextFromTextBody(textBody) : "";

                    slide.Shapes.Add(new ExtractedShape
                    {
                        Name = name,
                        Type = "GroupedShape",
                        Text = text
                    });
                }
            }
        }

        // Images with hashes
        slide.Images = ExtractImages(slidePart);

        return slide;
    }

    private static string GetShapeType(P.Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<PlaceholderShape>();
        if (ph != null)
        {
            if (ph.Type?.HasValue == true)
            {
                var val = ph.Type.Value;
                if (val == PlaceholderValues.Title) return "Title";
                if (val == PlaceholderValues.CenteredTitle) return "CenteredTitle";
                if (val == PlaceholderValues.SubTitle) return "Subtitle";
                if (val == PlaceholderValues.Body) return "Body";
                if (val == PlaceholderValues.Object) return "Content";
                if (val == PlaceholderValues.DateAndTime) return "DateAndTime";
                if (val == PlaceholderValues.Footer) return "Footer";
                if (val == PlaceholderValues.SlideNumber) return "SlideNumber";
                if (val == PlaceholderValues.Header) return "Header";
                return val.ToString();
            }
            return "Placeholder";
        }
        return "TextBox";
    }

    private static string GetTextFromTextBody(OpenXmlCompositeElement textBody)
    {
        var paragraphs = textBody.Elements<A.Paragraph>();
        return string.Join("\n", paragraphs.Select(p =>
            string.Join("", p.Elements<A.Run>().Select(r => r.Text?.Text ?? ""))
        ));
    }

    private static string GetSlideLayoutName(SlidePart slidePart)
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

    private static string? GetSlideNotes(SlidePart slidePart)
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
                    var paragraphs = textBody.Elements<A.Paragraph>();
                    string text = string.Join("\n", paragraphs.Select(p =>
                        string.Join("", p.Elements<A.Run>().Select(r => r.Text?.Text ?? ""))
                    ));
                    return string.IsNullOrWhiteSpace(text) ? null : text;
                }
            }
        }
        return null;
    }

    private static List<string> GetSlideComments(SlidePart slidePart, PresentationPart presPart)
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
            comments.Add($"{authorName}: {text}");
        }

        return comments;
    }

    private static List<ExtractedImage> ExtractImages(SlidePart slidePart)
    {
        var images = new List<ExtractedImage>();

        foreach (var imagePart in slidePart.ImageParts)
        {
            var info = new ExtractedImage
            {
                ContentType = imagePart.ContentType,
                FileName = Path.GetFileName(imagePart.Uri.ToString()),
            };

            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();
            info.SizeBytes = bytes.Length;
            info.Sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

            images.Add(info);
        }

        return images;
    }
}
