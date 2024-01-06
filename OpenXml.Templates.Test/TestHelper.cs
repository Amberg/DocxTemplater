using System.Diagnostics;

namespace OpenXml.Templates.Test
{
    internal static class TestHelper
    {
        public static void SaveAsFileAndOpenInWord(this Stream stream)
        {
#if !DEBUG
return;
#endif

            stream.Position = 0;
            var fileName = Path.ChangeExtension(Path.GetTempFileName(), "docx");
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
            proc?.WaitForExit();
        }
    }
}
