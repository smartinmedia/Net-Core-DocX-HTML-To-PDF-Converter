using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocXToPdfConverter
{
    
    public class ReplacementDictionaries
    {
        public string NewLineTag { get; set; }
        //Start and End Tags can e. g. be both "<###>"
        public string TextReplacementStartTag { get; set; }
        public string TextReplacementEndTag { get; set; }
        public Dictionary<string,string> TextReplacements { get; set; }

        /*
         * Important: The MemoryStream may carry an image.
         * Allowed file types: JPEG/JPG, BMP, TIFF, GIF, PNG
         */

        //Take different replacement tags here, else there may be collision with the text replacements,
        //e. g. "<+++>" 
        public string ImageReplacementStartTag { get; set; }
        public string ImageReplacementEndTag { get; set; }

        public Dictionary<string, MemoryStream> JpegReplacements { get; set; }
    }
}
