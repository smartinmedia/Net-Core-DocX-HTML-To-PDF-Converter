using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocXToPdfConverter;
using OpenXmlPowerTools;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


namespace Website.BackgroundWorkers
{
    public class DocXHandler

    {
        private MemoryStream _docxMs;
        private ReplacementDictionaries _rep;

        public DocXHandler(string docXTemplateFilename, ReplacementDictionaries rep)
        {
            _docxMs = StreamHandler.GetFileAsMemoryStream(docXTemplateFilename);
            _rep = rep;
            
        }

        public void ReplaceTextsAndImages()
        {
            ReplaceTexts();
            
        }
       

        public MemoryStream ReplaceTexts()
        {
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(_docxMs, true))
            {
                //REMOVE THESE Markups, because they break up the text into multiple pieces, 
                //thereby preventing simple search and replace
                SimplifyMarkupSettings settings = new SimplifyMarkupSettings
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
                    ReplaceTabsWithSpaces = true
                };
                MarkupSimplifier.SimplifyMarkup(doc, settings);

                var document = doc.MainDocumentPart.Document;

                foreach (var text in document.Descendants<Text>()) // <<< Here
                {
                    foreach (var replace in _rep.TextReplacements)
                    {
                        if (text.Text.Contains(_rep.TextReplacementStartTag + replace.Key + _rep.TextReplacementEndTag))
                        {
                            if (replace.Value.Contains(_rep.NewLineTag))//If we have line breaks present
                            {
                                string[] repArray = replace.Value.Split(new string[] {_rep.NewLineTag}, StringSplitOptions.None);

                                var lastInsertedText = text;
                                var lastInsertedBreak = new Break();

                                for (var i = 0; i < repArray.Length; i++)
                                {
                                    if (i == 0)//The text is only replaced with the first part of the replacement array
                                    {
                                        text.Text = text.Text.Replace(_rep.TextReplacementStartTag + replace.Key + _rep.TextReplacementEndTag, repArray[i]);

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
                                text.Text = text.Text.Replace(_rep.TextReplacementStartTag + replace.Key + _rep.TextReplacementEndTag, replace.Value);

                            }
                        }

                    }
                }

            }

            _docxMs.Position = 0;
            return _docxMs;
        }

        public void AppendImageToElement(MemoryStream imageMemoryStream)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(_docxMs, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(GetImagePartTypeFromMemStream(imageMemoryStream));

                imagePart.FeedData(imageMemoryStream);
                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        }

        private ImagePartType GetImagePartTypeFromMemStream(MemoryStream stream)
        {
            var image = Image.FromStream(stream);

            if (ImageFormat.Jpeg.Equals(image.RawFormat))
            {
                return ImagePartType.Jpeg;
            }
            else if (ImageFormat.Png.Equals(image.RawFormat))
            {
                return ImagePartType.Png;
            }
            else if (ImageFormat.Gif.Equals(image.RawFormat))
            {
                return ImagePartType.Gif;
            }
            else if (ImageFormat.Bmp.Equals(image.RawFormat))
            {
                return ImagePartType.Bmp;
            }
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
            {
                return ImagePartType.Tiff;
            }

            return ImagePartType.Jpeg;
        }

        
    }
}
