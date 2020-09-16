using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using A = DocumentFormat.OpenXml.Drawing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;


namespace DocXToPdfConverter.DocXToPdfHandlers
{
    public class DocXHandler
    {
        private readonly MemoryStream _docxMs;
        private readonly Placeholders _rep;
        private int _imageCounter;

        public DocXHandler(string docXTemplateFilename, Placeholders rep)
        {
            _docxMs = StreamHandler.GetFileAsMemoryStream(docXTemplateFilename);
            _rep = rep;

        }


        public MemoryStream ReplaceAll()
        {
            if (_rep != null)
            {
                ReplaceTexts();

                ReplaceHyperlinks();

                ReplaceTableRows();

                ReplaceImages();
            }

            _docxMs.Position = 0;

            return _docxMs;
        }


        public MemoryStream ReplaceTexts()
        {
            if (_rep.TextPlaceholders == null || _rep.TextPlaceholders.Count == 0)
                return null;

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                // Search in body, headers and footers
                var documentTexts = doc.MainDocumentPart.Document.Descendants<Text>();
                var headerTexts = doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Text>());
                var footerTexts = doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Text>());

                foreach (var text in documentTexts.Concat(headerTexts).Concat(footerTexts)) // <<< Here
                {
                    foreach (var replace in _rep.TextPlaceholders)
                    {
                        if (text.Text.Contains(_rep.TextPlaceholderStartTag + replace.Key + _rep.TextPlaceholderEndTag))
                        {
                            if (!string.IsNullOrEmpty(replace.Value) && replace.Value.Contains(_rep.NewLineTag))//If we have line breaks present
                            {
                                string[] repArray = replace.Value.Split(new string[] { _rep.NewLineTag }, StringSplitOptions.None);

                                var lastInsertedText = text;
                                Break lastInsertedBreak;

                                for (var i = 0; i < repArray.Length; i++)
                                {
                                    if (i == 0)//The text is only replaced with the first part of the replacement array
                                    {
                                        text.Text = text.Text.Replace(_rep.TextPlaceholderStartTag + replace.Key + _rep.TextPlaceholderEndTag, repArray[i]);

                                    }
                                    else
                                    {
                                        var tmpText = new Text(repArray[i]);
                                        var tmpBreak = new Break();
                                        text.Parent.InsertAfter(tmpBreak, lastInsertedText);
                                        lastInsertedBreak = tmpBreak;
                                        text.Parent.InsertAfter(tmpText, lastInsertedBreak);
                                        lastInsertedText = tmpText;
                                    }

                                }

                            }
                            else
                            {
                                text.Text = text.Text.Replace(_rep.TextPlaceholderStartTag + replace.Key + _rep.TextPlaceholderEndTag, replace.Value ?? string.Empty);

                            }
                        }

                    }
                }

            }

            _docxMs.Position = 0;
            return _docxMs;
        }

        public MemoryStream ReplaceHyperlinks()
        {
            if (_rep.HyperlinkPlaceholders == null || _rep.HyperlinkPlaceholders.Count == 0)
                return null;

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                // Search in body, headers and footers
                var documentTexts = doc.MainDocumentPart.Document.Descendants<Text>();

                foreach (var text in documentTexts)
                {
                    foreach (var replace in _rep.HyperlinkPlaceholders)
                    {
                        var pl = _rep.HyperlinkPlaceholderStartTag + replace.Key + _rep.HyperlinkPlaceholderEndTag;
                        if (text.Text.Contains(pl))
                        {
                            var run = text.Ancestors<Run>().First();

                            if (text.Text.StartsWith(pl))
                            {
                                var newAfterRun = (Run)run.Clone();
                                string afterText = text.Text.Substring(pl.Length, text.Text.Length - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;

                                run.Parent.InsertAfter(newAfterRun, run);
                            }
                            else if (text.Text.EndsWith(pl))
                            {
                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, text.Text.Length - pl.Length);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;

                                run.Parent.InsertBefore(newBeforeRun, run);
                            }
                            else
                            {
                                //Break the texts into the part before and after image. Then create separate runs for them
                                var pos = text.Text.IndexOf(pl, StringComparison.CurrentCulture);

                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, pos);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;
                                run.Parent.InsertBefore(newBeforeRun, run);

                                var newAfterRun = (Run)run.Clone();
                                string afterText = text.Text.Substring(pos + pl.Length, text.Text.Length - pos - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;
                                run.Parent.InsertAfter(newAfterRun, run);
                            }

                            var relation = doc.MainDocumentPart.AddHyperlinkRelationship(new Uri(replace.Value.Link, UriKind.RelativeOrAbsolute), true);
                            string relationid = relation.Id;
                            var linkText = string.IsNullOrEmpty(replace.Value.Text) ? replace.Value.Link : replace.Value.Text;
                            var hyper =
                                new Hyperlink(
                                    new Run(
                                        new RunProperties(new RunStyle() { Val = "Hyperlink" }),
                                        new Text(linkText)))
                                {
                                    Id = relationid,
                                    History = OnOffValue.FromBoolean(true)
                                };

                            run.Parent.InsertBefore(hyper, run);
                            run.Remove();
                        }
                    }
                }
            }

            _docxMs.Position = 0;
            return _docxMs;
        }

        public MemoryStream ReplaceTableRows()
        {
            if (_rep.TablePlaceholders == null || _rep.TablePlaceholders.Count == 0)
                return null;

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                foreach (var trDict in _rep.TablePlaceholders) //Take a Row (one Dictionary) at a time
                {
                    var trCol0 = trDict.First();
                    // Find the text elements matching one of the search strings
                    // where the text is inside a table cell --> these are the rows we are searching for.
                    var textElements =
                        doc.MainDocumentPart.Document.Body.Descendants<Text>()
                        .Concat(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Text>()))
                        .Concat(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Text>()))
                        .Where(t =>
                            trDict.Keys.Select(key => _rep.TablePlaceholderStartTag + trCol0.Key + _rep.TablePlaceholderEndTag).Any(s => t.Text.Contains(s)) &&
                            t.Ancestors<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Any());

                    // Loop through all found rows
                    foreach (var textElement in textElements)
                    {
                        var newTableRows = new List<TableRow>();
                        var tableRow = textElement.Ancestors<TableRow>().First();

                        for (var j = 0; j < trCol0.Value.Length; j++) //Lets create row by row and replace placeholders
                        {
                            newTableRows.Add((TableRow)tableRow.CloneNode(true));
                            var tableRowCopy = newTableRows[newTableRows.Count - 1];

                            foreach (var text in tableRow.Descendants<Text>()) //Cycle through the cells of the row to replace from the Dictionary value ( string array)
                            {
                                for (var index = 0;
                                    index < trDict.Count;
                                    index++) //Now cycle through the "columns" (keys) of the Dictionary and replace item by item
                                {
                                    var item = trDict.ElementAt(index);

                                    if (text.Text.Contains(_rep.TablePlaceholderStartTag + item.Key + _rep.TablePlaceholderEndTag))
                                    {
                                        if (item.Value[j].Contains(_rep.NewLineTag)) //If we have line breaks present
                                        {
                                            string[] repArray = item.Value[j].Split(new string[] { _rep.NewLineTag },
                                                StringSplitOptions.None);

                                            var lastInsertedText = text;
                                            Break lastInsertedBreak;

                                            for (var i = 0; i < repArray.Length; i++)
                                            {
                                                if (i == 0
                                                ) //The text is only replaced with the first part of the replacement array
                                                {
                                                    text.Text = text.Text.Replace(
                                                        _rep.TablePlaceholderStartTag + item.Key +
                                                        _rep.TablePlaceholderEndTag, repArray[i]);
                                                }
                                                else
                                                {
                                                    var tmpText = new Text(repArray[i]);
                                                    var tmpBreak = new Break();
                                                    text.Parent.InsertAfter(tmpBreak, lastInsertedText);
                                                    lastInsertedBreak = tmpBreak;
                                                    text.Parent.InsertAfter(tmpText, lastInsertedBreak);
                                                    lastInsertedText = tmpText;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            text.Text = text.Text.Replace(
                                                _rep.TablePlaceholderStartTag + item.Key + _rep.TablePlaceholderEndTag,
                                                item.Value[j]);
                                        }
                                    }
                                }
                            }

                            tableRow.Parent.InsertAfter(tableRowCopy, tableRow);
                            tableRow = tableRowCopy;
                        }

                        tableRow.Remove();
                    }

                }

            }
            _docxMs.Position = 0;
            return _docxMs;
        }


        public MemoryStream ReplaceImages()
        {
            if (_rep.ImagePlaceholders == null || _rep.ImagePlaceholders.Count == 0)
                return null;

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                var documentTexts = doc.MainDocumentPart.Document.Descendants<Text>();

                foreach (var text in documentTexts)
                {
                    foreach (var replace in _rep.ImagePlaceholders)
                    {
                        string pl = _rep.ImagePlaceholderStartTag + replace.Key + _rep.ImagePlaceholderEndTag;
                        _imageCounter++;
                        if (text.Text.Contains(pl))
                        {
                            var run = text.Ancestors<Run>().First();
                            var newRunForImage = new Run();
                            //Break the texts into the part before and after image. Then create separate runs for them
                            var pos = text.Text.IndexOf(pl, StringComparison.CurrentCulture);

                            if (pos == 0)
                            {
                                var newAfterRun = (Run)run.Clone();
                                string afterText = text.Text.Substring(pl.Length, text.Text.Length - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;

                                run.Parent.InsertAfter(newAfterRun, run);
                            }
                            else if (text.Text.EndsWith(pl))
                            {
                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, pos);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;

                                run.Parent.InsertBefore(newBeforeRun, run);
                            }
                            else
                            {
                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, pos);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;
                                run.Parent.InsertBefore(newBeforeRun, run);

                                var newAfterRun = (Run)run.Clone();
                                string afterText = text.Text.Substring(pos + pl.Length, text.Text.Length - pos - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;
                                run.Parent.InsertAfter(newAfterRun, run);
                            }

                            run.Parent.InsertBefore(newRunForImage, run);
                            run.Remove();

                            AppendImageToElement(replace, newRunForImage, doc);
                        }

                    }
                }
            }
            _docxMs.Position = 0;
            return _docxMs;

        }


        private void AppendImageToElement(KeyValuePair<string, ImageElement> placeholder, OpenXmlElement element, WordprocessingDocument wordprocessingDocument)
        {
            string imageExtension = placeholder.Value.MemStream.GetImageType();

            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

            var imageUri = new Uri($"/word/media/{placeholder.Key}{_imageCounter}.{imageExtension}", UriKind.Relative);

            // Create "image" part in /word/media
            // Change content type for other image types.
            PackagePart packageImagePart = wordprocessingDocument.Package.CreatePart(imageUri, "Image/" + imageExtension);

            // Feed data.
            placeholder.Value.MemStream.Position = 0;
            byte[] imageBytes = placeholder.Value.MemStream.ToArray();
            packageImagePart.GetStream().Write(imageBytes, 0, imageBytes.Length);

            PackagePart documentPackagePart = mainPart.OpenXmlPackage.Package.GetPart(new Uri("/word/document.xml", UriKind.Relative));

            // URI to the image is relative to relationship document.
            PackageRelationship imageRelationshipPart = documentPackagePart.CreateRelationship(
                new Uri("media/" + placeholder.Key + _imageCounter + "." + imageExtension, UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            var imgTmp = placeholder.Value.MemStream.GetImage();

            var drawing = GetImageElement(imageRelationshipPart.Id, placeholder.Key, "picture", imgTmp.Width, imgTmp.Height, placeholder.Value.Dpi);
            element.AppendChild(drawing);
        }



        private Drawing GetImageElement(
            string imagePartId,
            string fileName,
            string pictureName,
            double width,
            double height,
            double ppi)
        {
            double englishMetricUnitsPerInch = 914400;
            double pixelsPerInch = ppi;

            //calculate size in emu
            double emuWidth = width * englishMetricUnitsPerInch / pixelsPerInch;
            double emuHeight = height * englishMetricUnitsPerInch / pixelsPerInch;

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (UInt32Value)1U, Name = pictureName + _imageCounter },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = fileName },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                    {
                                        Embed = imagePartId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                            new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle })))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });
            return element;
        }

        private static void CleanMarkup(WordprocessingDocument doc)
        {
            //REMOVE THESE Markups, because they break up the text into multiple pieces, 
            //thereby preventing simple search and replace
            var settings = new SimplifyMarkupSettings
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                ReplaceTabsWithSpaces = true,
                RemoveBookmarks = true
            };
            MarkupSimplifier.SimplifyMarkup(doc, settings);
        }
    }
}
