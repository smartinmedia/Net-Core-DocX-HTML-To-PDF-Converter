using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace DocXToPdfConverter
{
    public static class ImageHandler
    {

        public static Image GetImageFromStream(MemoryStream ms)
        {
            var image = System.Drawing.Image.FromStream(ms);
            ms.Position = 0;
            return image;
        }

        public static ImagePartType GetImagePartTypeFromMemStream(MemoryStream stream)
        {
            stream.Position = 0;
            var image = Image.FromStream(stream);
            stream.Position = 0;


            if (ImageFormat.Jpeg.Equals(image.RawFormat))
            {
                return ImagePartType.Jpeg;
            }
            else if (ImageFormat.Png.Equals(image.RawFormat))
            {
                return ImagePartType.Png;
            }
            else if (ImageFormat.Gif.Equals(image.RawFormat))
            {
                return ImagePartType.Gif;
            }
            else if (ImageFormat.Bmp.Equals(image.RawFormat))
            {
                return ImagePartType.Bmp;
            }
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
            {
                return ImagePartType.Tiff;
            }

            return ImagePartType.Jpeg;
        }
    }
}
