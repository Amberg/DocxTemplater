#if DEBUG
using System.Diagnostics;
#endif

namespace DocxTemplater.Test
{
    internal static class TestHelper
    {
        public static void SaveAsFileAndOpenInWord(this Stream stream, string extension = "docx")
        {
#if DEBUG
            stream.Position = 0;
            var fileName = Path.ChangeExtension(Path.GetTempFileName(), extension);
            using (var fileStream = File.OpenWrite(fileName))
            {
                stream.CopyTo(fileStream);
            }

            ProcessStartInfo psi = new()
            {
                FileName = fileName,
                UseShellExecute = true
            };
            using var proc = Process.Start(psi);
#endif
        }
    }
}
