using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace DocXToPdfConverter
{
    //THIS ALL COMES FROM: https://github.com/Reflexe/doc_to_pdf

    public class LibreOfficeFailedException : Exception
    {
        public LibreOfficeFailedException(int exitCode)
            : base(string.Format("LibreOffice has failed with " + exitCode))
        { }
    }

    public static class ConvertDocxToPdfWithLibreOffice
    {


        private static string GetLibreOfficePath()
        {
            switch (Environment.OSVersion.Platform)
            {
                case PlatformID.Unix:
                    return "/usr/bin/soffice";
                case PlatformID.Win32NT:
                    string binaryDirectory = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    return binaryDirectory + "\\Windows\\program\\soffice.exe";
                default:
                    throw new PlatformNotSupportedException("Your OS is not supported");
            }
        }



        //libreOfficePath for Windows: e. g. "C:\\program\\soffice.exe


        //With Portable Apps it is here: C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe

        public static void ConvertToPdf(string docxFile, string pdfFile, string libreOfficePath)
        {
            if (libreOfficePath == "")
            {
                libreOfficePath = GetLibreOfficePath();
            }
            
            ProcessStartInfo procStartInfo =
                new ProcessStartInfo(libreOfficePath, String.Format("--convert-to pdf --nologo --headless --outdir {0} {1}", System.IO.Path.GetDirectoryName(pdfFile), docxFile));
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            procStartInfo.CreateNoWindow = true;
            procStartInfo.WorkingDirectory = Environment.CurrentDirectory;

            Process process = new Process() {StartInfo = procStartInfo,};
            process.Start();
            process.WaitForExit();

            // Check for failed exit code.
            if (process.ExitCode != 0)
            {
                throw new LibreOfficeFailedException(process.ExitCode);
            }
            else
            {
                System.IO.File.Move(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pdfFile), System.IO.Path.GetFileNameWithoutExtension(docxFile)+".pdf"), pdfFile);
            }

        }




    }




}
