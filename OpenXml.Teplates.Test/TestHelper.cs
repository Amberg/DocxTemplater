using System.Diagnostics;

namespace OpenXml.Templates.Test
{
    internal static class TestHelper
    {
        public static void SaveAsFileAndOpenInWord(this Stream stream)
        {
            stream.Position = 0;
            var fileName = Path.ChangeExtension(Path.GetTempFileName(), "docx");
            using (var fileStream = File.OpenWrite(fileName))
            {
                stream.CopyTo(fileStream);
            }

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = fileName;
            psi.UseShellExecute = true;
            using var proc = Process.Start(psi);
            proc.WaitForExit();
        }
    }
}
