using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxLetterGenerator
{
    public static class DocxConverter
    {
        public static void ReplacePlaceholders(string docxPath, Dictionary<string, string> replacements)
        {
            if (!File.Exists(docxPath))
                throw new FileNotFoundException($"Файл не найден: {docxPath}");

            if (replacements == null || replacements.Count == 0)
                return;

            using (WordprocessingDocument document = WordprocessingDocument.Open(docxPath, true))
            {
                if (document.MainDocumentPart == null)
                    throw new Exception("В документе нет основной части");

                Body body = document.MainDocumentPart.Document.Body;
                if (body != null)
                {
                    foreach (var paragraph in body.Descendants<Paragraph>())
                        ReplaceInParagraph(paragraph, replacements);
                }

                foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                {
                    if (headerPart.Header != null)
                    {
                        foreach (var paragraph in headerPart.Header.Descendants<Paragraph>())
                            ReplaceInParagraph(paragraph, replacements);
                        headerPart.Header.Save();
                    }
                }

                foreach (var footerPart in document.MainDocumentPart.FooterParts)
                {
                    if (footerPart.Footer != null)
                    {
                        foreach (var paragraph in footerPart.Footer.Descendants<Paragraph>())
                            ReplaceInParagraph(paragraph, replacements);
                        footerPart.Footer.Save();
                    }
                }

                document.MainDocumentPart.Document.Save();
            }
        }

        private static void ReplaceInParagraph(Paragraph paragraph, Dictionary<string, string> replacements)
        {
            List<Text> textElements = paragraph.Descendants<Text>().ToList();
            if (textElements.Count == 0) return;

            string fullText = string.Concat(textElements.Select(t => t.Text));
            bool changed = false;

            foreach (var pair in replacements)
            {
                if (fullText.Contains(pair.Key))
                {
                    fullText = fullText.Replace(pair.Key, pair.Value ?? string.Empty);
                    changed = true;
                }
            }

            if (changed)
            {
                textElements[0].Text = fullText;
                for (int i = 1; i < textElements.Count; i++)
                    textElements[i].Text = string.Empty;
            }
        }

        public static List<Paragraph> BuildAppendixListBlocks(List<AppendixItem> appendices)
        {
            var blocks = new List<Paragraph>();

            if (appendices == null || appendices.Count == 0)
                return blocks;

            blocks.Add(CreateParagraph("Приложения:", JustificationValues.Left, true));

            for (int i = 0; i < appendices.Count; i++)
            {
                var app = appendices[i];
                string pagesText = string.IsNullOrWhiteSpace(app.Pages) ? "" : $" на {app.Pages} л.";
                blocks.Add(CreateParagraph($"{i + 1}. {app.Title}{pagesText}", JustificationValues.Left, false));
            }

            return blocks;
        }

        public static List<Paragraph> BuildAppendixBlocks(List<AppendixItem> appendices)
        {
            var blocks = new List<Paragraph>();

            if (appendices == null || appendices.Count == 0)
                return blocks;

            for (int i = 0; i < appendices.Count; i++)
            {
                var app = appendices[i];

                if (i > 0)
                {
                    var pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                    blocks.Add(pageBreak);
                }

                string header = appendices.Count == 1 ? "Приложение" : $"Приложение {i + 1}";
                blocks.Add(CreateParagraph(header, JustificationValues.Left, false));
                blocks.Add(CreateParagraph(app.Title, JustificationValues.Center, true));

                if (!string.IsNullOrWhiteSpace(app.Text))
                {
                    var lines = app.Text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                    foreach (var line in lines)
                        blocks.Add(CreateParagraph(line, JustificationValues.Left, false));
                }
            }

            return blocks;
        }

        public static void ReplacePlaceholderWithBlocks(string docxPath, string placeholder, List<Paragraph> blocks)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(docxPath, true))
            {
                if (document.MainDocumentPart?.Document?.Body == null) return;

                Paragraph targetParagraph = document.MainDocumentPart.Document.Body
                    .Descendants<Paragraph>()
                    .FirstOrDefault(p => GetParagraphText(p).Contains(placeholder));

                if (targetParagraph == null) return;

                var current = targetParagraph;
                foreach (var block in blocks)
                {
                    var inserted = current.InsertAfterSelf(block.CloneNode(true));
                    current = (Paragraph)inserted;
                }

                targetParagraph.Remove();
                document.MainDocumentPart.Document.Save();
            }
        }

        public static void RemovePlaceholder(string docxPath, string placeholder)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(docxPath, true))
            {
                if (document.MainDocumentPart?.Document?.Body == null) return;

                Paragraph targetParagraph = document.MainDocumentPart.Document.Body
                    .Descendants<Paragraph>()
                    .FirstOrDefault(p => GetParagraphText(p).Contains(placeholder));

                if (targetParagraph != null)
                {
                    targetParagraph.Remove();
                    document.MainDocumentPart.Document.Save();
                }
            }
        }

        private static string GetParagraphText(Paragraph paragraph)
        {
            return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
        }

        private static Paragraph CreateParagraph(string text, JustificationValues alignment, bool bold)
        {
            var paragraph = new Paragraph();
            paragraph.Append(new ParagraphProperties(new Justification() { Val = alignment }));

            var run = new Run();
            if (bold)
                run.RunProperties = new RunProperties(new Bold());

            run.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);

            return paragraph;
        }
    }
}