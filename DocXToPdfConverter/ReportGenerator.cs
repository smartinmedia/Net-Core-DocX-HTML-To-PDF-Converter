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


using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocXToPdfConverter.DocXToPdfHandlers;


 
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

        public void Convert(string inputFile, string outputFile, Placeholders rep = null)
        {
            if (inputFile.EndsWith(".docx"))
            {
                if (outputFile.EndsWith(".docx"))
                {
                    GenerateReportFromDocxToDocX(inputFile, outputFile, rep);
                }
                else if (outputFile.EndsWith(".pdf"))
                {
                    GenerateReportFromDocxToPdf(inputFile, outputFile, rep);
                }
                else if(outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
                {
                    GenerateReportFromDocxToHtml(inputFile, outputFile, rep);
                }
            }
            else if (inputFile.EndsWith(".html") || inputFile.EndsWith(".htm"))
            {
                if (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
                {
                    GenerateReportFromHtmlToHtml(inputFile, outputFile, rep);
                }
                else if (outputFile.EndsWith(".docx"))
                {
                    
                    GenerateReportFromHtmlToDocx(inputFile, outputFile, rep);
                }
                else if (outputFile.EndsWith(".pdf"))
                {
                    GenerateReportFromHtmlToPdf(inputFile, outputFile, rep);
                }
            }
        }

        //string docxSource = filename with path
        private void GenerateReportFromDocxToDocX(string docxSource, string docxTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            StreamHandler.WriteMemoryStreamToDisk(ms, docxTarget);
        }

        ////string docxSource = filename with path
        private void GenerateReportFromDocxToPdf(string docxSource, string pdfTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.Combine(Path.GetDirectoryName(pdfTarget), Path.GetFileNameWithoutExtension(pdfTarget) + Guid.NewGuid().ToString().Substring(0,10)+".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            ConvertWithLibreOffice.Convert(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }

        
        private void GenerateReportFromDocxToHtml(string docxSource, string htmlTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.Combine(Path.GetDirectoryName(htmlTarget), Path.GetFileNameWithoutExtension(docxSource)+Guid.NewGuid().ToString().Substring(0,10) + ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            ConvertWithLibreOffice.Convert(tmpFile, htmlTarget, _locationOfLibreOfficeSoffice);
            //PtConvertDocxToHtml.ConvertToHtml(tmpFile, htmlTargetDirectory);
            File.Delete(tmpFile);

        }


        private void GenerateReportFromHtmlToHtml(string htmlSource, string htmlTarget, Placeholders rep)
        {
            string html = File.ReadAllText(htmlSource);
            html = HtmlHandler.ReplaceAll(html, rep);
            File.WriteAllText(htmlTarget, html);
        }


        //string htmlSource = filename to a *.html/*.htm file with path
        private void GenerateReportFromHtmlToDocx(string htmlSource, string docxTarget, Placeholders rep)
        {
            var tmpFile = Path.Combine(Path.GetDirectoryName(docxTarget), Path.GetFileNameWithoutExtension(htmlSource) + Guid.NewGuid().ToString().Substring(0, 10) + ".html");
            GenerateReportFromHtmlToHtml(htmlSource, tmpFile, rep);
            ConvertWithLibreOffice.Convert(tmpFile, docxTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);

        }


        //This requires the HtmlAgilityPack
        //string htmlSource = filename to a *.html/*.htm file with path
        private void GenerateReportFromHtmlToPdf(string htmlSource, string pdfTarget, Placeholders rep)
        {
            var tmpFile = Path.Combine(Path.GetDirectoryName(pdfTarget), Path.GetFileNameWithoutExtension(htmlSource) + Guid.NewGuid().ToString().Substring(0, 10) + ".html");
            GenerateReportFromHtmlToHtml(htmlSource, tmpFile, rep);
            ConvertWithLibreOffice.Convert(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);

        }


    }
}
