using System;
using System.Collections.Generic;
using System.Text;
using SkiaSharp;

namespace DocXToPdfConverter
{
    public static class FontFamily
    {
        public static IEnumerable<string> GetFontFamilies()
        {
            using (var fontManager = SKFontManager.Default)
            {
                return fontManager.FontFamilies;
            }
        }
    }
}
