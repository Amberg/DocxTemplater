using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Xml;

namespace DocxTemplater.Images
{
    internal class SvgHelper
    {
        private const int FallbackDim = 300;

        internal static string CreateSvgPart<T>(T parent, byte[] svgBytes)
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(ImagePartType.Svg);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new MemoryStream(svgBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }

        internal static bool TryReadAsSvg(
            byte[] bytes,
            out int width,
            out int height)
        {
            width = height = FallbackDim;

            // --- 1) ultra-cheap reject: must start with '<' (skip BOM + whitespace) ----
            int i = (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) ? 3 : 0;
            while (i < bytes.Length && bytes[i] <= 0x20)
            {
                i++;
            }

            if (i >= bytes.Length || bytes[i] != (byte)'<')
            {
                return false;
            }

            // --- 2) small, real parse with XmlReader ----
            try
            {
                using var ms = new MemoryStream(bytes, writable: false);
                var settings = new XmlReaderSettings
                {
                    ConformanceLevel = ConformanceLevel.Document,
                    DtdProcessing = DtdProcessing.Prohibit
                };
                using var xr = XmlReader.Create(ms, settings);

                xr.MoveToContent();
                if (!xr.Name.Equals("svg", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                int? parsedWidth = ParseDim(xr.GetAttribute("width"));
                int? parsedHeight = ParseDim(xr.GetAttribute("height"));

                // fall back to viewBox if one dim is missing
                if ((parsedWidth is null || parsedHeight is null) && xr.GetAttribute("viewBox") is { } vb)
                {
                    var p = vb.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    if (p.Length == 4)
                    {
                        parsedWidth ??= ParseDim(p[2]);
                        parsedHeight ??= ParseDim(p[3]);
                    }
                }

                width = parsedWidth ?? FallbackDim;
                height = parsedHeight ?? FallbackDim;
                return true;
            }
            catch (XmlException)
            {
                return false;
            }

            static int? ParseDim(string raw)
            {
                return int.TryParse(raw?.TrimEnd('%', 'p', 'x'), out var v) ? v : null;
            }
        }
    }
}