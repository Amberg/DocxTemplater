using System.Diagnostics;


namespace DocxTemplater.Test
{
    internal static class TestHelper
    {
        public static void SaveAsFileAndOpenInWord(this Stream stream, string extension = "docx")
        {
#if DEBUG
            return;
#pragma warning disable CS0162
#endif
            if (Environment.GetEnvironmentVariable("DOCX_TEMPLATER_VISUAL_TESTING") == null)
            {
                return;
            }
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
#pragma warning restore CS0162
        }
    }
}
