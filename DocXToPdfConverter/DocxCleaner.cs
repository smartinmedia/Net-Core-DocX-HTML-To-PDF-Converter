using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace DocXToPdfConverter
{
    /*
     *
     *  D  E  P  R  A  C  A  T  E  D
     *
     *
     *
     */


    public static class DocxCleaner
    {

        public static MemoryStream Clean2(MemoryStream _docxMs)
        {
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(_docxMs, true))
            {

                var document = doc.MainDocumentPart.Document;

                foreach (var text in document.Descendants<Text>()) // <<< Here
                {

                }
            }

            return _docxMs;
        }



        public static MemoryStream Clean(MemoryStream ms)
        {
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(ms, true))
            {
                XDocument xDoc = doc.MainDocumentPart.GetXDocument();
                CleanUp(xDoc);
                doc.MainDocumentPart.SaveXDocument();

                /*
                foreach (var h in doc.MainDocumentPart.HeaderParts)
                {
                    xDoc = h.GetXDocument();
                    CleanUp(xDoc);
                    h.SaveXDocument();
                }
                */
                foreach (var f in doc.MainDocumentPart.FooterParts)
                {
                    xDoc = f.GetXDocument();
                    CleanUp(xDoc);
                    f.SaveXDocument();
                }
            }

            return ms;
        }

        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader streamReader = new StreamReader(part.GetStream()))
                xdoc = XDocument.Load(XmlReader.Create(streamReader));
            part.AddAnnotation(xdoc);
            return xdoc;
        }

        public static void SaveXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
            {
                using (XmlWriter xw =
                  XmlWriter.Create(part.GetStream(FileMode.Create, FileAccess.Write)))
                    xdoc.WriteTo(xw);
            }
        }
    


        // get rid of every rsid attribute/element in the doc.
        // they exist to enable merging of forked documents; not something
        // we're interested in here.  if we don't delete these nodes, they
        // show up as changed.
        private static void CleanUp(XDocument doc)
        {
            XNamespace w =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            doc.Descendants().Attributes(w + "rsidTr").Remove();
            doc.Descendants().Attributes(w + "rsidSect").Remove();
            doc.Descendants().Attributes(w + "rsidRDefault").Remove();
            doc.Descendants().Attributes(w + "rsidR").Remove();
            doc.Descendants().Attributes(w + "rsidDel").Remove();
            doc.Descendants().Attributes(w + "rsidP").Remove();
            doc.Descendants().Attributes(w + "lang").Remove();

            doc.Descendants(w + "lang").Remove();
            doc.Descendants(w + "rsid").Remove();
        }

        
    }





}

