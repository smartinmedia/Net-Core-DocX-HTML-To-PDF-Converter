/* Net-Core-DocX-HTML-To-PDF-Converter
 * https://github.com/smartinmedia/Net-Core-DocX-HTML-To-PDF-Converter
 *
 *
 * This application was coded (c) by Dr. Martin Weihrauch 2019
 * for Smart In Media GmbH & Co / https://www.smartinmedia.com
 * DISTRIBUTED UNDER THE MIT LICENSE
 *
 *
 */


using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocXToPdfConverter;

namespace ExampleApplication
{
    class Program
    {
        static void Main(string[] args)
        {

            //This is only to get this example to work (find the word docx and the html file, which were
            //shipped with this.
            string executableLocation = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string docxLocation = Path.Combine(executableLocation, "Test-Template.docx");
            string htmlLocation = Path.Combine(executableLocation, "Test-HTML-page.html");

            /*
             *
             * You have to create Placeholder objects from the Placeholders.cs class
             * Only create, what you need.
             * Here, we create one object for a docx conversion example and one
             * for an HTML conversion example.
             *
             *  DOCX OBJECT
             *
             */


            //Prepare texts, which you want to insert into the custom fields in the template (remember
            //to use start and stop tags.
            //NOTE that line breaks can be inserted as what you define them in ReplacementDictionaries.NewLineTag (here we use <br/>).

            var placeholders = new Placeholders();
            placeholders.NewLineTag = "<br/>";
            placeholders.TextPlaceholderStartTag = "##";
            placeholders.TextPlaceholderEndTag = "##";
            placeholders.TablePlaceholderStartTag = "==";
            placeholders.TablePlaceholderEndTag = "==";
            placeholders.ImagePlaceholderStartTag = "++";
            placeholders.ImagePlaceholderEndTag = "++";

            //You should be able to also use other OpenXML tags in your strings
            placeholders.TextPlaceholders = new Dictionary<string, string>
            {
                {"Name", "Mr. Miller" },
                {"Street", "89 Brook St" },
                {"City", "Brookline MA 02115<br/>USA" },
                {"InvoiceNo", "5" },
                {"Total", "U$ 4,500" },
                {"Date", "28 Jul 2019" }
            };



            //Table ROW replacements are a little bit more complicated: With them you can
            //fill out only one table row in a table and it will add as many rows as you 
            //need, depending on the string Array.
            placeholders.TablePlaceholders = new List<Dictionary<string, string[]>>
            {
                
                    new Dictionary<string, string[]>()
                    {
                        {"Name", new string[]{ "Homer Simpson", "Mr. Burns", "Mr. Smithers" }},
                        {"Department", new string[]{ "Power Plant", "Administration", "Administration" }},
                        {"Responsibility", new string[]{ "Oversight", "CEO", "Assistant" }},
                        {"Telephone number", new string[]{ "888-234-2353", "888-295-8383", "888-848-2803" }}
                    },
                    new Dictionary<string, string[]>()
                    {
                        {"Qty", new string[]{ "2", "5", "7" }},
                        {"Product", new string[]{ "Software development", "Customization", "Travel expenses" }},
                        {"Price", new string[]{ "U$ 2,000", "U$ 1,000", "U$ 1,500" }},
                    }

            };

            //You have to add the images as a memory stream to the Dictionary! Place a key (placeholder) into the docx template.
            //There is a method to read files as memory streams (GetFileAsMemoryStream)
            //We already did that with <+++>ProductImage<+++>

            var productImage =
                StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "ProductImage.jpg"));

            var qrImage =
                StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "QRCode.PNG"));

            placeholders.ImagePlaceholders = new Dictionary<string, MemoryStream>
            {
                {"QRCode", qrImage },
                {"ProductImage", productImage }
            };

            
            /*
             * Execution of conversion tests
             *
             *
             */

            //Most important: give the full path to the soffice.exe file including soffice.exe.
            //Don't know how that would be named on Linux...
            var test = new ReportGenerator(@"F:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe");


            test.GenerateReportFromHtmlToHtml(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.html"), placeholders);

            //test.GenerateReportFromHtmlToPdf("f:\\vmc\\simhtml.htm", "f:\\vmc\\simhtml.pdf", myDictionary);
            //test.GenerateReportFromDocxToDocX("Test-Template.docx", "F:\\vmc\\template.docx", myDictionary);
            //test.GenerateReportFromDocxToHtml("Test-Template.docx", "F:\\vmc\\template.html", placeholders);



            //test.GenerateReportFromDocxToPdf("Test-Template.docx", "F:\\vmc\\template.pdf", myDictionary);
        }
    }
}
