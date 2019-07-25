using System.IO;


namespace DocXToPdfConverter
{
    public class StreamHandler
    {
        public static MemoryStream GetFileAsMemoryStream(string filename)
        {
            MemoryStream ms = new MemoryStream();
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
                file.CopyTo(ms);
            return ms;
        }

        public static void WriteMemoryStreamToDisk(MemoryStream ms, string filename)
        {
            using (FileStream file = new FileStream(filename, FileMode.Create, System.IO.FileAccess.Write))
                ms.CopyTo(file);
        }

    }
}
