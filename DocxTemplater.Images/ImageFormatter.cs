using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures; // http://schemas.openxmlformats.org/drawingml/2006/picture"

namespace DocxTemplater.Images
{
    public class ImageFormatter : IFormatter
    {
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
                using var image = Image.Load(imageBytes);
                var imagePartType = DetectPartTypeInfo(context.Placeholder, image.Metadata);
                var root = target.GetRoot();
                string impagepartRelationShipId = null;
                uint maxPropertyId = 0;
                if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
                {
                    maxPropertyId = openXmlPartRootElement.OpenXmlPart.GetMaxDocPropertyId();
                    if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                    {
                        impagepartRelationShipId = CreateImagePart(headerPart, imageBytes, imagePartType);
                    }
                    else if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                    {
                        impagepartRelationShipId = CreateImagePart(footerPart, imageBytes, imagePartType);
                    }
                    else if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                    {
                        impagepartRelationShipId = CreateImagePart(mainDocumentPart, imageBytes, imagePartType);
                    }
                }

                if (impagepartRelationShipId == null)
                {
                    throw new OpenXmlTemplateException("Could not find a valid image part");
                }

                // case 1. Image ist the only child element of a <wps:wsp> (TextBox)
                if (TryHandleImageInWordprocessingShape(target, impagepartRelationShipId, image,
                        context.Args.FirstOrDefault(), maxPropertyId))
                {
                    return;
                }

                AddInlineGraphicToRun(target, impagepartRelationShipId, image, maxPropertyId);
            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new OpenXmlTemplateException("Could not detect image format", e);
            }
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
        private static bool TryHandleImageInWordprocessingShape(Text target, string impagepartRelationShipId, Image image,
            string firstArgument, uint maxPropertyId)
        {
            var drawing = target.GetFirstAncestor<Drawing>();
            if (drawing == null)
            {
                return false;
            }

            var anchor = target.GetFirstAncestor<DW.Anchor>();
            if (anchor == null)
            {
                return false;
            }

            var targetExtent = anchor.GetFirstChild<DW.Extent>();
            if (targetExtent != null)
            {
                double scale = 0;
                var imageCx = image.Width * 9525;
                var imageCy = image.Height * 9525;
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

                ReplaceAnchorContentWithPicture(impagepartRelationShipId, maxPropertyId, drawing);
            }

            target.Remove();
            return true;
        }


        private static void ReplaceAnchorContentWithPicture(string impagepartRelationShipId, uint maxDocumentPropertyId, Drawing original)
        {
            var propertyId = maxDocumentPropertyId + 1;
            var originalAnchor = original.GetFirstChild<DW.Anchor>();
            var originaleExtent = originalAnchor.GetFirstChild<DW.Extent>();

            var horzPosition = originalAnchor.GetFirstChild<DW.HorizontalPosition>().CloneNode(true);
            var vertPosition = originalAnchor.GetFirstChild<DW.VerticalPosition>().CloneNode(true);

            var anchorChildElments = new OpenXmlElement[]
            {
                new DW.SimplePosition {X = 0L, Y = 0L},
                horzPosition,
                vertPosition,
                new DW.Extent {Cx = originaleExtent.Cx, Cy = originaleExtent.Cy},
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.WrapNone(),
                new DW.DocProperties
                {
                    Id = propertyId,
                    Name = $"Picture {propertyId}"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks {NoChangeAspect = true}),
                new A.Graphic(
                    new A.GraphicData(
                            CreatePicture(impagepartRelationShipId, propertyId, originaleExtent.Cx, originaleExtent.Cy)
                        )
                        {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            };

            var anchor = originalAnchor.CloneNode(false);
            anchor.Append(anchorChildElments);
            var dw = new Drawing(anchor);
            original.InsertAfterSelf(dw);
            original.Remove();
        }

        private static void AddInlineGraphicToRun(Text target, string impagepartRelationShipId, Image image,
            uint maxDocumentPropertyId)
        {
            var propertyId = maxDocumentPropertyId + 1;
            var cx = image.Width * 9525;
            var cy = image.Height * 9525;
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
                                    CreatePicture(impagepartRelationShipId, propertyId, cx, cy)
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

        private static PIC.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy)
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
                        new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(
                            new A.AdjustValueList()
                        )
                    { Preset = A.ShapeTypeValues.Rectangle }));
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