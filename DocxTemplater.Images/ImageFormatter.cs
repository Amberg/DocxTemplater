using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures; // http://schemas.openxmlformats.org/drawingml/2006/picture"

namespace DocxTemplater.Images
{
    public class ImageFormatter : IFormatter
    {
        private sealed record ImageInfo(int PixelWidth, int PixelHeight, string ImagePartRelationId);
        private readonly Dictionary<byte[], ImageInfo> m_imagePartRelIdCache = new();
        private OpenXmlPartRootElement m_currentRoot;

        public bool CanHandle(Type type, string prefix)
        {
            var prefixUpper = prefix.ToUpper();
            return prefixUpper is "IMAGE" or "IMG" && type == typeof(byte[]);
        }

        public void ApplyFormat(FormatterContext context, Text target)
        {
            // TODO: handle oter ppi values than default 96
            // see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pixelsperinch?view=openxml-2.8.1#remarks
            if (context.Value is not byte[] imageBytes)
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
                        using var image = Image.Load(imageBytes);
                        string imagePartRelId = null;
                        var imagePartType = DetectPartTypeInfo(context.Placeholder, image.Metadata);
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
                        imageInfo = new ImageInfo(image.Width, image.Height, imagePartRelId);
                        m_imagePartRelIdCache[imageBytes] = imageInfo;
                    }

                    // case 1. Image ist the only child element of a <wps:wsp> (TextBox)
                    if (TryHandleImageInWordprocessingShape(target, imageInfo, context.Args.FirstOrDefault() ?? string.Empty, maxPropertyId))
                    {
                        return;
                    }

                    AddInlineGraphicToRun(target, imageInfo, maxPropertyId);
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
                var imageCx = imageInfo.PixelWidth * 9525;
                var imageCy = imageInfo.PixelHeight * 9525;
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

                ReplaceAnchorContentWithPicture(imageInfo.ImagePartRelationId, maxPropertyId, drawing);
            }

            target.Remove();
            return true;
        }


        private static void ReplaceAnchorContentWithPicture(string impagepartRelationShipId, uint maxDocumentPropertyId,
            Drawing original)
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
                            CreatePicture(impagepartRelationShipId, propertyId, originaleExtent.Cx, originaleExtent.Cy, rotation)
                        )
                        {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            });
            var dw = new Drawing(clonedInlineOrAnchor);
            original.InsertAfterSelf(dw);
            original.Remove();
        }

        private static void AddInlineGraphicToRun(Text target, ImageInfo imageInfo, uint maxDocumentPropertyId)
        {
            var propertyId = maxDocumentPropertyId + 1;
            var cx = imageInfo.PixelWidth * 9525;
            var cy = imageInfo.PixelHeight * 9525;
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
                                    CreatePicture(imageInfo.ImagePartRelationId, propertyId, cx, cy, 0)
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

        private static PIC.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, int rotation)
        {
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
                        Rotation = rotation
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
    }
}