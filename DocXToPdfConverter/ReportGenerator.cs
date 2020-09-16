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
using System.IO;
using DocXToPdfConverter.DocXToPdfHandlers;


namespace DocXToPdfConverter
{
    public class ReportGenerator
    {
        private readonly string _locationOfLibreOfficeSoffice;

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
                else if (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
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


        /// <summary>
        /// Prints the file and optionally generates from placeholders
        /// </summary>
        /// <param name="templateFile">The input file. May be docx, html or pdf</param>
        /// <param name="printerName">optional printername to print on. If this value is empty, the default printer is used</param>
        /// <param name="rep">A collection of placeholders to be applied if the input file is docx or html.</param>
        public void Print(string templateFile, string printerName = null, Placeholders rep = null)
        {
            if (rep != null)
            {
                if (templateFile.EndsWith(".docx"))
                {
                    PrintDocx(templateFile, printerName, rep);
                }
                else if (templateFile.EndsWith(".html") || templateFile.EndsWith(".htm"))
                {
                    PrintHtml(templateFile, printerName, rep);
                }
                else
                {
                    LibreOfficeWrapper.Print(templateFile, printerName, _locationOfLibreOfficeSoffice);
                }
            }
            else
            {
                LibreOfficeWrapper.Print(templateFile, printerName, _locationOfLibreOfficeSoffice);
            }
        }


        private void PrintDocx(string templateFile, string printername, Placeholders rep)
        {
            var docx = new DocXHandler(templateFile, rep);
            var ms = docx.ReplaceAll();
            var tempFileToPrint = Path.ChangeExtension(Path.GetTempFileName(), ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tempFileToPrint);
            LibreOfficeWrapper.Print(tempFileToPrint, printername, _locationOfLibreOfficeSoffice);
            File.Delete(tempFileToPrint);
        }


        private void PrintHtml(string templateFile, string printername, Placeholders rep)
        {
            var htmlContent = File.ReadAllText(templateFile);
            htmlContent = HtmlHandler.ReplaceAll(htmlContent, rep);
            var tempFileToPrint = Path.ChangeExtension(Path.GetTempFileName(), ".html");
            File.WriteAllText(tempFileToPrint, htmlContent);
            LibreOfficeWrapper.Print(tempFileToPrint, printername, _locationOfLibreOfficeSoffice);
            File.Delete(tempFileToPrint);
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
            var tmpFile = Path.Combine(Path.GetDirectoryName(pdfTarget), Path.GetFileNameWithoutExtension(pdfTarget) + Guid.NewGuid().ToString().Substring(0, 10) + ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            LibreOfficeWrapper.Convert(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }


        private void GenerateReportFromDocxToHtml(string docxSource, string htmlTarget, Placeholders rep)
        {
            var docx = new DocXHandler(docxSource, rep);
            var ms = docx.ReplaceAll();
            var tmpFile = Path.Combine(Path.GetDirectoryName(htmlTarget), Path.GetFileNameWithoutExtension(docxSource) + Guid.NewGuid().ToString().Substring(0, 10) + ".docx");
            StreamHandler.WriteMemoryStreamToDisk(ms, tmpFile);
            LibreOfficeWrapper.Convert(tmpFile, htmlTarget, _locationOfLibreOfficeSoffice);
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
            LibreOfficeWrapper.Convert(tmpFile, docxTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }


        //This requires the HtmlAgilityPack
        //string htmlSource = filename to a *.html/*.htm file with path
        private void GenerateReportFromHtmlToPdf(string htmlSource, string pdfTarget, Placeholders rep)
        {
            var tmpFile = Path.Combine(Path.GetDirectoryName(pdfTarget), Path.GetFileNameWithoutExtension(htmlSource) + Guid.NewGuid().ToString().Substring(0, 10) + ".html");
            GenerateReportFromHtmlToHtml(htmlSource, tmpFile, rep);
            LibreOfficeWrapper.Convert(tmpFile, pdfTarget, _locationOfLibreOfficeSoffice);
            File.Delete(tmpFile);
        }


    }
}
