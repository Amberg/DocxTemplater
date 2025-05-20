using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTemplater.Images
{
    public class ImageFormatter : IFormatter
    {
        private static readonly Regex ArgumentRegex = new(@"(?<key>[whr]):(?<value>\d+)(?<unit>px|cm|in|pt)?", RegexOptions.Compiled, TimeSpan.FromMilliseconds(500));
        private sealed record ImageInfo(int PixelWidth, int PixelHeight, string ImagePartRelationId, int? Orientation, bool IsSvg = false);
        private readonly Dictionary<byte[], ImageInfo> m_imagePartRelIdCache = new();
        private OpenXmlPartRootElement m_currentRoot;

        public bool CanHandle(Type type, string prefix)
        {
            var prefixUpper = prefix.ToUpper();
            return prefixUpper is "IMAGE" or "IMG" && type == typeof(byte[]);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            // TODO: handle other ppi values than default 96
            // see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pixelsperinch?view=openxml-2.8.1#remarks
            if (formatterContext.Value is not byte[] imageBytes)
            {
                return;
            }
            if (imageBytes.Length == 0)
            {
                target.Text = string.Empty;
                return;
            }
            try
            {
                var root = target.GetRoot();
                if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
                {
                    var maxPropertyId = openXmlPartRootElement.OpenXmlPart.GetMaxDocPropertyId();

                    if (!TryGetImageIdFromCache(imageBytes, openXmlPartRootElement, out var imageInfo))
                    {
                        if (IsSvgImage(imageBytes))
                        {
                            // Handle SVG specifically
                            string imagePartRelId = null;
                            
                            if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                            {
                                imagePartRelId = CreateSvgPart(headerPart, imageBytes);
                            }
                            else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                            {
                                imagePartRelId = CreateSvgPart(footerPart, imageBytes);
                            }
                            else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                            {
                                imagePartRelId = CreateSvgPart(mainDocumentPart, imageBytes);
                            }
                            
                            if (imagePartRelId == null)
                            {
                                throw new OpenXmlTemplateException("Could not create SVG part");
                            }
                            
                            // Parse size arguments from SVG if possible or use defaults
                            int defaultWidth = ExtractSvgWidth(imageBytes) ?? 300;
                            int defaultHeight = ExtractSvgHeight(imageBytes) ?? 300;
                            
                            imageInfo = new ImageInfo(defaultWidth, defaultHeight, imagePartRelId, null, true);
                            m_imagePartRelIdCache[imageBytes] = imageInfo;
                        }
                        else
                        {
                            // Handle regular images with ImageSharp
                            using var image = Image.Load(imageBytes);
                            int? orientation = null;
                            if (image.Metadata?.ExifProfile?.TryGetValue(ExifTag.Orientation, out var orientationValue) == true)
                            {
                                orientation = orientationValue.Value;
                            }

                            string imagePartRelId = null;
                            var imagePartType = DetectPartTypeInfo(formatterContext.Placeholder, image.Metadata);
                            if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                            {
                                imagePartRelId = CreateImagePart(headerPart, imageBytes, imagePartType);
                            }
                            else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                            {
                                imagePartRelId = CreateImagePart(footerPart, imageBytes, imagePartType);
                            }
                            else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                            {
                                imagePartRelId = CreateImagePart(mainDocumentPart, imageBytes, imagePartType);
                            }
                            if (imagePartRelId == null)
                            {
                                throw new OpenXmlTemplateException("Could not find a valid image part");
                            }
                            imageInfo = new ImageInfo(image.Width, image.Height, imagePartRelId, orientation);
                            m_imagePartRelIdCache[imageBytes] = imageInfo;
                        }
                    }

                    if (TryHandleImageInWordprocessingShape(target, imageInfo, formatterContext.Args.FirstOrDefault() ?? string.Empty, maxPropertyId))
                    {
                        return;
                    }

                    AddInlineGraphicToRun(target, imageInfo, maxPropertyId, formatterContext.Args);
                }
                else
                {
                    throw new OpenXmlTemplateException("Could not find root to insert image");
                }
            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new OpenXmlTemplateException("Could not detect image format", e);
            }
        }

        private bool TryGetImageIdFromCache(byte[] imageBytes, OpenXmlPartRootElement root, out ImageInfo imageInfo)
        {
            if (m_currentRoot != root)
            {
                m_imagePartRelIdCache.Clear();
                m_currentRoot = root;
            }
            return m_imagePartRelIdCache.TryGetValue(imageBytes, out imageInfo);
        }

        private static bool IsSvgImage(byte[] imageBytes)
        {
            // Check for SVG XML signature at the beginning of the file
            try
            {
                // Check the first portion of the file for SVG signature
                string content = Encoding.UTF8.GetString(imageBytes, 0, Math.Min(imageBytes.Length, 1000)).Trim();
                return (content.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) || 
                        content.StartsWith("<svg", StringComparison.OrdinalIgnoreCase)) && 
                       content.Contains("<svg", StringComparison.OrdinalIgnoreCase) && 
                       content.Contains("xmlns", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static int? ExtractSvgWidth(byte[] svgBytes)
        {
            try
            {
                string svgContent = Encoding.UTF8.GetString(svgBytes);
                var widthMatch = Regex.Match(svgContent, @"<svg[^>]*\s+width\s*=\s*[""']?(\d+)(?:px)?[""']?");
                if (widthMatch.Success && int.TryParse(widthMatch.Groups[1].Value, out int width))
                {
                    return width;
                }
                
                // Check for viewBox as fallback
                var viewBoxMatch = Regex.Match(svgContent, @"<svg[^>]*\s+viewBox\s*=\s*[""']?(\d+)\s+(\d+)\s+(\d+)\s+(\d+)[""']?");
                if (viewBoxMatch.Success && int.TryParse(viewBoxMatch.Groups[3].Value, out int viewBoxWidth))
                {
                    return viewBoxWidth;
                }
            }
            catch
            {
                // Ignore parsing errors and use default
            }
            return null;
        }

        private static int? ExtractSvgHeight(byte[] svgBytes)
        {
            try
            {
                string svgContent = Encoding.UTF8.GetString(svgBytes);
                var heightMatch = Regex.Match(svgContent, @"<svg[^>]*\s+height\s*=\s*[""']?(\d+)(?:px)?[""']?");
                if (heightMatch.Success && int.TryParse(heightMatch.Groups[1].Value, out int height))
                {
                    return height;
                }
                
                // Check for viewBox as fallback
                var viewBoxMatch = Regex.Match(svgContent, @"<svg[^>]*\s+viewBox\s*=\s*[""']?(\d+)\s+(\d+)\s+(\d+)\s+(\d+)[""']?");
                if (viewBoxMatch.Success && int.TryParse(viewBoxMatch.Groups[4].Value, out int viewBoxHeight))
                {
                    return viewBoxHeight;
                }
            }
            catch
            {
                // Ignore parsing errors and use default
            }
            return null;
        }

        private static PartTypeInfo DetectPartTypeInfo(string modelPath, ImageMetadata imageMetadata)
        {
            return imageMetadata switch
            {
                { DecodedImageFormat.Name: "TIFF" } => ImagePartType.Tiff,
                { DecodedImageFormat.Name: "BMP" } => ImagePartType.Bmp,
                { DecodedImageFormat.Name: "GIF" } => ImagePartType.Gif,
                { DecodedImageFormat.Name: "JPEG" } => ImagePartType.Jpeg,
                { DecodedImageFormat.Name: "PNG" } => ImagePartType.Png,
                _ => throw new OpenXmlTemplateException($"Could not detect image format for image in {modelPath}")
            };
        }

        /// <summary>
        ///     If the image is contained in a "wsp" element (TextBox), the text box is used as a container for the image.
        ///     the size of the text box is adjusted to the size of the image.
        /// </summary>
        private static bool TryHandleImageInWordprocessingShape(Text target, ImageInfo imageInfo,
            string firstArgument, uint maxPropertyId)
        {
            var drawing = target.GetFirstAncestor<Drawing>();
            if (drawing == null)
            {
                return false;
            }

            // get extent of the drawing either from the anchor or inline
            var targetExtent = target.GetFirstAncestor<DW.Anchor>()?.GetFirstChild<DW.Extent>() ?? target.GetFirstAncestor<DW.Inline>()?.GetFirstChild<DW.Extent>();
            if (targetExtent != null)
            {
                double scale = 0;
                var imageCx = OpenXmlHelper.PixelsToEmu(imageInfo.PixelWidth);
                var imageCy = OpenXmlHelper.PixelsToEmu(imageInfo.PixelHeight);
                if (firstArgument.Equals("KEEPRATIO", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = Math.Min(targetExtent.Cx / (double)imageCx, targetExtent.Cy / (double)imageCy);
                }
                else if (firstArgument.Equals("STRETCHW", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = targetExtent.Cx / (double)imageCx;
                }
                else if (firstArgument.Equals("STRETCHH", StringComparison.CurrentCultureIgnoreCase))
                {
                    scale = targetExtent.Cy / (double)imageCy;
                }

                if (scale > 0)
                {
                    targetExtent.Cx = (long)(imageCx * scale);
                    targetExtent.Cy = (long)(imageCy * scale);
                }

                ReplaceAnchorContentWithPicture(imageInfo.ImagePartRelationId, maxPropertyId, drawing, imageInfo.Orientation);
            }

            target.Remove();
            return true;
        }


        private static void ReplaceAnchorContentWithPicture(string impagepartRelationShipId, uint maxDocumentPropertyId,
            Drawing original, int? originalOrientation)
        {
            var propertyId = maxDocumentPropertyId + 1;
            var inlineOrAnchor = (OpenXmlElement)original.GetFirstChild<DW.Anchor>() ??
                                 (OpenXmlElement)original.GetFirstChild<DW.Inline>();
            var originaleExtent = inlineOrAnchor.GetFirstChild<DW.Extent>();
            var transform = inlineOrAnchor.Descendants<A.Transform2D>().FirstOrDefault();
            int rotation = transform?.Rotation ?? 0;
            var clonedInlineOrAnchor = inlineOrAnchor.CloneNode(false);

            if (inlineOrAnchor is DW.Anchor anchor)
            {
                clonedInlineOrAnchor.Append(new DW.SimplePosition { X = 0L, Y = 0L });
                var horzPosition = anchor.GetFirstChild<DW.HorizontalPosition>().CloneNode(true);
                var vertPosition = inlineOrAnchor.GetFirstChild<DW.VerticalPosition>().CloneNode(true);
                clonedInlineOrAnchor.Append(horzPosition);
                clonedInlineOrAnchor.Append(vertPosition);
                clonedInlineOrAnchor.Append(new DW.Extent { Cx = originaleExtent.Cx, Cy = originaleExtent.Cy });
                clonedInlineOrAnchor.Append(new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                });
                clonedInlineOrAnchor.Append(new DW.WrapNone());
            }
            else if (inlineOrAnchor is DW.Inline)
            {
                clonedInlineOrAnchor.Append(new DW.Extent { Cx = originaleExtent.Cx, Cy = originaleExtent.Cy });
                clonedInlineOrAnchor.Append(new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                });
            }

#pragma warning disable IDE0300
            clonedInlineOrAnchor.Append(new OpenXmlElement[]
            {
                new DW.DocProperties
                {
                    Id = propertyId,
                    Name = $"Picture {propertyId}"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks {NoChangeAspect = true}),
                new A.Graphic(
                    new A.GraphicData(
                            CreatePicture(impagepartRelationShipId, propertyId, originaleExtent.Cx, originaleExtent.Cy, rotation, originalOrientation)
                        )
                        {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            });
            var dw = new Drawing(clonedInlineOrAnchor);
            original.InsertAfterSelf(dw);
            original.Remove();
        }

        /// <summary>
        /// If image is not part of a textbox this method is used to add the image to the run.
        /// </summary>
        private static void AddInlineGraphicToRun(Text target, ImageInfo imageInfo, uint maxDocumentPropertyId, string[] arguments)
        {
            var propertyId = maxDocumentPropertyId + 1;

            TransformSize(imageInfo.PixelWidth, imageInfo.PixelHeight, arguments, out var cx, out var cy, out var rotation);

            // Define the reference of the image.
            var drawing =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = cx, Cy = cy },
                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties
                        {
                            Id = propertyId,
                            Name = $"Picture {propertyId}"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                    CreatePicture(imageInfo.ImagePartRelationId, propertyId, cx, cy, rotation, imageInfo.Orientation)
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U
                    });

            target.InsertAfterSelf(drawing);
            target.Remove();
        }


        /// <summary>
        /// Transforms the width and height of the image to the size in EMU.
        /// arguments is in format:
        /// Examples:
        /// "w:90px;h:90px;r:90"
        /// "w:90cm;h:90cm;r:90"
        /// "h:90cm;r:90"
        /// "w:90cm;h:90cm"
        /// "w:90cm;h:90cm;r:90"
        /// available units are px, cm, in, pt
        /// </summary>
        private static void TransformSize(int pixelWidth, int pixelHeight, string[] arguments, out int outCxEmu, out int outCyEmu, out int rot)
        {
            var cxEmu = -1;
            var cyEmu = -1;
            var rotation = 0;

            if (arguments == null || arguments.Length == 0)
            {
                outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                rot = rotation;
                return;
            }

            foreach (var argument in arguments)
            {
                try
                {
                    var matches = ArgumentRegex.Matches(argument);
                    if (matches.Count == 0)
                    {
                        outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                        outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                        rot = rotation;
                        return;
                    }

                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        var key = match.Groups["key"].Value;
                        var value = int.Parse(match.Groups["value"].Value);
                        var unit = match.Groups["unit"].Value;
                        switch (key)
                        {
                            case "w":
                                cxEmu = OpenXmlHelper.LengthToEmu(value, unit);
                                break;
                            case "h":
                                cyEmu = OpenXmlHelper.LengthToEmu(value, unit);
                                break;
                            case "r":
                                rotation = value;
                                break;
                        }
                    }
                }
                catch (RegexMatchTimeoutException)
                {
                    throw new OpenXmlTemplateException($"Invalid image formatter argument '{argument}'");
                }
            }

            if (cxEmu == -1 && cyEmu == -1)
            {
                outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
                rot = rotation;
                return;
            }

            if (cxEmu == -1)
            {
                cxEmu = (int)(cyEmu * ((double)pixelWidth / pixelHeight));
            }
            else if (cyEmu == -1)
            {
                cyEmu = (int)(cxEmu * ((double)pixelHeight / pixelWidth));
            }
            else
            {
                var aspectRatio = (double)pixelWidth / pixelHeight;
                var newAspectRatio = (double)cxEmu / cyEmu;
                if (aspectRatio > newAspectRatio)
                {
                    cyEmu = (int)(cxEmu / aspectRatio);
                }
                else
                {
                    cxEmu = (int)(cyEmu * aspectRatio);
                }
            }
            outCxEmu = cxEmu;
            outCyEmu = cyEmu;
            rot = rotation;
        }

        private static PIC.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, int rotation, int? originalOrientation = null)
        {
            int finalRotation = rotation;
            if (originalOrientation.HasValue)
            {
                switch (originalOrientation.Value)
                {
                    case 6: // Rotated 90 degrees right
                        finalRotation = (rotation + 90) % 360;
                        break;
                    case 3: // Rotated 180 degrees
                        finalRotation = (rotation + 180) % 360;
                        break;
                    case 8: // Rotated 90 degrees left
                        finalRotation = (rotation + 270) % 360;
                        break;
                }
            }

            finalRotation *= 60000; // Convert to OpenXML rotation units

            return new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties
                    {
                        Id = (UInt32Value)0U,
                        Name = $"Image{propertyId}.jpg"
                    },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip(new A.BlipExtensionList(
                        new A.BlipExtension
                        {
                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                        })
                    )
                    {
                        Embed = impagepartRelationShipId,
                        CompressionState = A.BlipCompressionValues.Print
                    },
                    new A.Stretch(
                        new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = cx, Cy = cy })
                    {
                        Rotation = finalRotation
                    },
                    new A.PresetGeometry(
                            new A.AdjustValueList()
                        )
                    {
                        Preset = A.ShapeTypeValues.Rectangle
                    }));
        }

        private static string CreateImagePart<T>(T parent, byte[] imageBytes, PartTypeInfo partType)
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(partType);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new MemoryStream(imageBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }

        private static string CreateSvgPart<T>(T parent, byte[] svgBytes)
            where T : OpenXmlPartContainer, ISupportedRelationship<ImagePart>
        {
            var imagePart = parent.AddImagePart(ImagePartType.Svg);
            var relationshipId = parent.GetIdOfPart(imagePart);
            var memStream = new MemoryStream(svgBytes);
            imagePart.FeedData(memStream);
            return relationshipId;
        }
    }
}