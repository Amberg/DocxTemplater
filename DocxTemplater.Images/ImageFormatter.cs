using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using System;
using System.Text.RegularExpressions;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Linq;

namespace DocxTemplater.Images
{
    public class ImageFormatter : IFormatter, IImageServiceProvider
    {
        private static readonly Regex ArgumentRegex = new(@"(?<key>[whr]):(?<value>\d+)(?<unit>px|cm|in|pt|mm)?", RegexOptions.Compiled, TimeSpan.FromMilliseconds(500));

        public bool CanHandle(Type type, string prefix)
        {
            var prefixUpper = prefix.ToUpper();
            return prefixUpper is "IMAGE" or "IMG" && type == typeof(byte[]);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext, Text target)
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
                var maxPropertyId = templateContext.ImageService.GetImage(root, imageBytes, out ImageInformation imageInfo);
                // Image ist a child element of a <wps:wsp> (TextBox)
                if (!TryHandleImageInWordprocessingShape(target, imageInfo, formatterContext.Args.FirstOrDefault() ?? string.Empty, maxPropertyId, templateContext.ImageService))
                {// Image is not a child element of a <wps:wsp> (TextBox) - rotation and scale is determined by the arguments
                    AddInlineGraphicToRun(target, imageInfo, maxPropertyId, formatterContext.Args, templateContext.ImageService);
                }

            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new OpenXmlTemplateException("Could not detect image format", e);
            }
        }

        /// <summary>
        /// If image is not part of a textbox this method is used to add the image to the run.
        /// </summary>
        private static void AddInlineGraphicToRun(Text target, ImageInformation imageInfo, uint maxDocumentPropertyId, string[] arguments, IImageService imageService)
        {
            var propertyId = maxDocumentPropertyId + 1;

            TransformSize(imageInfo.PixelWidth, imageInfo.PixelHeight, arguments, out var cx, out var cy, out var rotation);
            rotation = rotation.AddUnits(imageInfo.ExifRotation.Units);

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
                                    imageService.CreatePicture(imageInfo.ImagePartRelationId, propertyId, cx, cy, rotation)
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

        private static bool TryHandleImageInWordprocessingShape(Text target, ImageInformation imageInfo,
            string firstArgument, uint maxPropertyId, IImageService imageService)
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

                ReplaceAnchorContentWithPicture(imageInfo.ImagePartRelationId, maxPropertyId, drawing, imageInfo.ExifRotation, imageService);
            }

            target.Remove();
            return true;
        }

        private static void ReplaceAnchorContentWithPicture(string impagepartRelationShipId, uint maxDocumentPropertyId, Drawing original, ImageRotation imageInfoExifRotation, IImageService imageService)
        {
            var propertyId = maxDocumentPropertyId + 1;
            var inlineOrAnchor = (OpenXmlElement)original.GetFirstChild<DW.Anchor>() ??
                                 (OpenXmlElement)original.GetFirstChild<DW.Inline>();
            var originaleExtent = inlineOrAnchor.GetFirstChild<DW.Extent>();
            var transform = inlineOrAnchor.Descendants<A.Transform2D>().FirstOrDefault();
            var rotation = imageInfoExifRotation.AddUnits(transform?.Rotation ?? 0);
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
                            imageService.CreatePicture(impagepartRelationShipId, propertyId, originaleExtent.Cx, originaleExtent.Cy, rotation)
                        )
                        {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            });
            var dw = new Drawing(clonedInlineOrAnchor);
            original.InsertAfterSelf(dw);
            original.Remove();
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
        private static void TransformSize(int pixelWidth, int pixelHeight, string[] arguments, out int outCxEmu, out int outCyEmu, out ImageRotation rotationInDegree)
        {
            var cxEmu = -1;
            var cyEmu = -1;
            rotationInDegree = ImageRotation.CreateFromDegree(0);

            if (arguments == null || arguments.Length == 0)
            {
                outCxEmu = OpenXmlHelper.PixelsToEmu(pixelWidth);
                outCyEmu = OpenXmlHelper.PixelsToEmu(pixelHeight);
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
                                rotationInDegree = ImageRotation.CreateFromDegree(value);
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
                // if both are set, the aspect ratio is kept
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
        }

        public IImageService CreateImageService()
        {
            return new ImageService();
        }
    }
}