using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocXToPdfConverter.DocXToPdfHandlers;
using HtmlAgilityPack;

/*
 * (c) by Smart In Media GmbH & Co. KG
 * www.smartinmedia.com
 * Distributed under the MIT License
 * Smart In Media does not give any warranty and cannot be held
 * liable for this code.
 *
 *
 *
 *
 */
 
namespace DocXToPdfConverter
{
    public class ReportGenerator
    {
        private string _locationOfLibreOfficeSoffice;

        // If you dont need conversion to PDF, you can leave the LocationOfLibreOfficeSoffice empty 
        // For Windows users: this must point to the ".exe" file, so \Path\Path\soffice.exe
        public ReportGenerator(string locationOfLibreOfficeSoffice = "")
        {
            _locationOfLibreOfficeSoffice = locationOfLibreOfficeSoffice;
        }

        //string docxSource = filename with path
        public void GenerateReportFromDocxToDocX(string docxSource, string docxTarget, ReplacementDictionaries rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            StreamHandler.WriteMemoryStreamToDisk(ms, docxTarget);
        }

        ////string docxSource = filename with path
        public void GenerateReportFromDocxToPdf(string docxSource, string pdfTarget, ReplacementDictionaries rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.GetFileNameWithoutExtension(pdfTarget) + ".tmp";
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            ConvertDocxToPdfWithLibreOffice.ConvertToPdf(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }

        //Please note that this is not a target file, but a target directory!
        public void GenerateReportFromDocxToHtml(string docxSource, string htmlTargetDirectory, ReplacementDictionaries rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.GetFileNameWithoutExtension(docxSource);
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);

            PtConvertDocxToHtml.ConvertToHtml(tmpFile, htmlTargetDirectory);
            File.Delete(tmpFile);

        }

        //This requires the HtmlAgilityPack
        //string htmlSource = filename to a *.html/*.htm file with path
        public void GenerateReportFromHtmlToDocx(string htmlSource, string pdfTarget, ReplacementDictionaries rep)
        {
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlSource);
            htmlDoc.OptionFixNestedTags = true;
            htmlDoc.OptionOutputAsXml = true;
            htmlDoc.OptionCheckSyntax = true;
            HtmlNode bodyNode = htmlDoc.DocumentNode;
            htmlSource = bodyNode.WriteTo();

            
        }


        //This requires the HtmlAgilityPack
        //string htmlSource = filename to a *.html/*.htm file with path
        public void GenerateReportFromHtmlToPdf(string htmlSource, string pdfTarget, ReplacementDictionaries rep)
        {

        }

        /*******************************************************************************************#
         *
         * FROM OPENXML POWERTOOLS
         *
         * HtmlToWmlConverter expects the HTML to be passed as an XElement, i.e. as XML.  While the HTML test files that
         * are included in Open-Xml-PowerTools are able to be read as XML, most HTML is not able to be read as XML.
         * The best solution is to use the HtmlAgilityPack, which can parse HTML and save as XML.  The HtmlAgilityPack
         * is licensed under the Ms-PL (same as Open-Xml-PowerTools) so it is convenient to include it in your solution,
         * and thereby you can convert HTML to XML that can be processed by the HtmlToWmlConverter.
         * 
         * A convenient way to get the DLL that has been checked out with HtmlToWmlConverter is to clone the repo at
         * https://github.com/EricWhiteDev/HtmlAgilityPack
         * 
         * That repo contains only the DLL that has been checked out with HtmlToWmlConverter.
         * 
         * Of course, you can also get the HtmlAgilityPack source and compile it to get the DLL.  You can find it at
         * http://codeplex.com/HtmlAgilityPack
         * 
         * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The example files
         * in this module do not require HtmlAgilityPack to run.
        *******************************************************************************************/
    }
}
