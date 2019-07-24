using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocXToPdfConverter
{
    
    public class ReplacementDictionaries
    {
        public string ReplacementStartTag = "<###>";
        public string ReplacementEndTag = "<###>";
        public Dictionary<string,string> TextReplacements { get; set; }

        /*
         * Important: The MemoryStream may carry an image.
         * Allowed file types: JPEG/JPG, BMP, TIFF, GIF, PNG
         */

        public Dictionary<string, MemoryStream> JpegReplacements { get; set; }
    }
}
