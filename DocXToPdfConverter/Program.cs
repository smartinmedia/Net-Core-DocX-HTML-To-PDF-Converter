using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Website.BackgroundWorkers;

namespace DocXToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {

            //Do you know the path to your word-template? Then you can omit this
            string executableLocation = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string xslLocation = Path.Combine(executableLocation, "Test-Template.docx");

            //Prepare texts, which you want to insert into the custom fields in the template (remember
            //to use start and stop tags and store in a Dictionary.
            var myDictionary = new ReplacementDictionaries();
            myDictionary.NewLineTag = "<br/>";
            myDictionary.TextReplacementStartTag = "##";
            myDictionary.TextReplacementEndTag = "##";
            myDictionary.ImageReplacementStartTag = "++";
            myDictionary.ImageReplacementEndTag = "++";

            //Note that line breaks can be inserted as <w:br />
            //You should be able to also use other OpenXML tags in your strings
            myDictionary.TextReplacements = new Dictionary<string, string>
            {
                {"Name", "Mr. Miller" },
                {"Street", "89 Brook St" },
                {"City", "Brookline MA 02115" },
                {"InvoiceNo", "5" },
                {"Qty", "2<br/>6<br/>4" },
                {"Product", "Software development<br/>Customization<br/>Travel expenses" },
                {"Price", "U$ 1,500<br/>U$ 2,500<br/>U$ 500" },
                {"Total", "U$ 4,500" }
            };

            //You have to add the image as a memory stream to the Dictionary! Place a key (placeholder) into the docx template
            //We already did that with <+++>ProductImage<+++>

            var productImage =
                StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "ProductImage.jpg"));

            var qrImage =
                StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "QRCode.PNG"));

            var doc = new DocXHandler(xslLocation, myDictionary);
            var docxStream = doc.ReplaceTexts();
            StreamHandler.WriteMemoryStreamToDisk(docxStream, "F:\\vmc\\out.docx");
        }
    }
}
