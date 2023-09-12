using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using IronSoftware.Drawing;
using static IronSoftware.Drawing.AnyBitmap;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    public static class ImageHandler
    {
        public static AnyBitmap GetImage(this MemoryStream ms)
        {
            ms.Position = 0;
            var image = AnyBitmap.FromStream(ms);
            ms.Position = 0;
            return image;
        }

        public static ImagePartType GetImagePartType(this MemoryStream stream)
        {
            stream.Position = 0;
            using (var image = AnyBitmap.FromStream(stream))
            {
                stream.Position = 0;


                if (ImageFormat.Jpeg.Equals(image.GetImageFormat()))
                {
                    return ImagePartType.Jpeg;
                }
                else if (ImageFormat.Png.Equals(image.GetImageFormat()))
                {
                    return ImagePartType.Png;
                }
                else if (ImageFormat.Gif.Equals(image.GetImageFormat()))
                {
                    return ImagePartType.Gif;
                }
                else if (ImageFormat.Bmp.Equals(image.GetImageFormat()))
                {
                    return ImagePartType.Bmp;
                }
                else if (ImageFormat.Tiff.Equals(image.GetImageFormat()))
                {
                    return ImagePartType.Tiff;
                }

                return ImagePartType.Jpeg;
            }
        }

        public static string GetImageType(this MemoryStream stream)
        {
            stream.Position = 0;
            using (var image = AnyBitmap.FromStream(stream))
            {
                stream.Position = 0;

                if (ImageFormat.Jpeg.Equals(image.GetImageFormat()))
                {
                    return "jpeg";
                }
                else if (ImageFormat.Png.Equals(image.GetImageFormat()))
                {
                    return "png";
                }
                else if (ImageFormat.Gif.Equals(image.GetImageFormat()))
                {
                    return "gif";
                }
                else if (ImageFormat.Bmp.Equals(image.GetImageFormat()))
                {
                    return "bmp";
                }
                else if (ImageFormat.Tiff.Equals(image.GetImageFormat()))
                {
                    return "tiff";
                }

                return "";
            }
        }

        public static string GetBase64(this MemoryStream stream)
        {
            byte[] imageBytes = stream.ToArray();

            // Convert byte[] to Base64 String
            return Convert.ToBase64String(imageBytes);
        }
    }
}
