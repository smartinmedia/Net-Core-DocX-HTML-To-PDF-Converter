using System.IO;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    public static class StreamHandler
    {
        public static MemoryStream GetFileAsMemoryStream(string filename)
        {
            MemoryStream ms = new MemoryStream();
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
                file.CopyTo(ms);
            ms.Position = 0;
            return ms;
        }

        public static void WriteMemoryStreamToDisk(MemoryStream ms, string filename)
        {
            ms.Position = 0;

            using (FileStream file = new FileStream(filename, FileMode.Create, System.IO.FileAccess.Write))
                ms.CopyTo(file);
        }

    }
}
