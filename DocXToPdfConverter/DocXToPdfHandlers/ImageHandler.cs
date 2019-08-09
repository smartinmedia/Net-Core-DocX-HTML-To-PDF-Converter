using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    public static class ImageHandler
    {

        public static Image GetImageFromStream(MemoryStream ms)
        {
            ms.Position = 0;
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

        public static string GetImageTypeFromMemStream(MemoryStream stream)
        {
            stream.Position = 0;
            var image = Image.FromStream(stream);
            stream.Position = 0;


            if (ImageFormat.Jpeg.Equals(image.RawFormat))
            {
                return "jpeg";
            }
            else if (ImageFormat.Png.Equals(image.RawFormat))
            {
                return "png";
            }
            else if (ImageFormat.Gif.Equals(image.RawFormat))
            {
                return "gif";
            }
            else if (ImageFormat.Bmp.Equals(image.RawFormat))
            {
                return "bmp";
            }
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
            {
                return "tiff";
            }

            return "";
        }

        public static string GetBase64FromMemStream(MemoryStream stream)
        {
            byte[] imageBytes = stream.ToArray();

            // Convert byte[] to Base64 String
            return Convert.ToBase64String(imageBytes);
        }
    }
}
