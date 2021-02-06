using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    //THIS ALL COMES FROM: https://github.com/Reflexe/doc_to_pdf

    //And very helpful: https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not 


    public class LibreOfficeFailedException : Exception
    {
        public LibreOfficeFailedException(int exitCode)
            : base(string.Format("LibreOffice has failed with " + exitCode))
        { }
    }

    public static class LibreOfficeWrapper
    {

        private static string GetLibreOfficePath()
        {
            switch (Environment.OSVersion.Platform)
            {
                case PlatformID.Unix:
                    return "/usr/bin/soffice";
                case PlatformID.Win32NT:
                    string binaryDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    return binaryDirectory + "\\Windows\\program\\soffice.exe";
                default:
                    throw new PlatformNotSupportedException("Your OS is not supported");
            }
        }



        //libreOfficePath for Windows: e. g. "C:\\program\\soffice.exe


        //With Portable Apps it is here: C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe

        public static void Convert(string inputFile, string outputFile, string libreOfficePath)
        {
            var commandArgs = new List<string>();
            string convertedFile = "";

            if (string.IsNullOrEmpty(libreOfficePath))
            {
                libreOfficePath = GetLibreOfficePath();
            }

            //Create tmp folder
            var tmpFolder = Path.Combine(Path.GetDirectoryName(outputFile), "DocXHtmlToPdfConverterTmp" + Guid.NewGuid().ToString().Substring(0, 10));
            if (!Directory.Exists(tmpFolder))
            {
                Directory.CreateDirectory(tmpFolder);
            }

            commandArgs.Add("--convert-to");

            if ((inputFile.EndsWith(".html") || inputFile.EndsWith(".htm")) && outputFile.EndsWith(".pdf"))
            {
                commandArgs.Add("pdf:writer_pdf_Export");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".pdf");
            }
            else if (inputFile.EndsWith(".docx") && outputFile.EndsWith(".pdf"))
            {
                commandArgs.Add("pdf:writer_pdf_Export");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".pdf");
            }
            else if (inputFile.EndsWith(".docx") && (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm")))
            {
                commandArgs.Add("html:HTML:EmbedImages");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".html");
            }
            else if ((inputFile.EndsWith(".html") || inputFile.EndsWith(".htm")) && (outputFile.EndsWith(".docx")))
            {
                commandArgs.Add("docx:\"Office Open XML Text\"");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".docx");
            }

            commandArgs.AddRange(new[] { inputFile, "--norestore", "--writer", "--headless", "--outdir", tmpFolder });

            var procStartInfo = new ProcessStartInfo(libreOfficePath);
            foreach (var arg in commandArgs) { procStartInfo.ArgumentList.Add(arg); }
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            procStartInfo.CreateNoWindow = true;
            procStartInfo.WorkingDirectory = Environment.CurrentDirectory;

            var process = new Process() { StartInfo = procStartInfo };
            Process[] pname = Process.GetProcessesByName("soffice");

            //Supposedly, only one instance of Libre Office can be run simultaneously
            while (pname.Length > 0)
            {
                Thread.Sleep(5000);
                pname = Process.GetProcessesByName("soffice");
            }

            process.Start();
            process.WaitForExit();

            // Check for failed exit code.
            if (process.ExitCode != 0)
            {
                throw new LibreOfficeFailedException(process.ExitCode);
            }
            else
            {
                if (File.Exists(outputFile)) File.Delete(outputFile);
                if (File.Exists(convertedFile))
                {
                    File.Move(convertedFile, outputFile);
                }
                ClearDirectory(tmpFolder);
                Directory.Delete(tmpFolder);
            }
        }


        public static void Print(string inputFile, string printerName, string libreOfficePath)
        {
            var commandArgs = new List<string>();

            if (string.IsNullOrEmpty(libreOfficePath))
            {
                libreOfficePath = GetLibreOfficePath();
            }

            commandArgs.Add("-p");

            if (!string.IsNullOrEmpty(printerName))
            {
                commandArgs.Add("-pt");
                commandArgs.Add(printerName);
            }

            commandArgs.AddRange(new[] { inputFile, "--norestore", "--writer", "--headless" });

            var procStartInfo = new ProcessStartInfo(libreOfficePath);
            foreach (var arg in commandArgs) { procStartInfo.ArgumentList.Add(arg); }
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            procStartInfo.CreateNoWindow = true;
            procStartInfo.WorkingDirectory = Environment.CurrentDirectory;

            var process = new Process() { StartInfo = procStartInfo };
            Process[] pname = Process.GetProcessesByName("soffice");

            //Supposedly, only one instance of Libre Office can be run simultaneously
            while (pname.Length > 0)
            {
                Thread.Sleep(5000);
                pname = Process.GetProcessesByName("soffice");
            }

            process.Start();
            process.WaitForExit();

            // Check for failed exit code.
            if (process.ExitCode != 0)
            {
                throw new LibreOfficeFailedException(process.ExitCode);
            }
        }


        private static void ClearDirectory(string folderName)
        {
            var dir = new DirectoryInfo(folderName);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            foreach (DirectoryInfo di in dir.GetDirectories())
            {
                ClearDirectory(di.FullName);
                di.Delete();
            }
        }

    }

}
