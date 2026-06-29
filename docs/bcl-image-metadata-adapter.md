# Dependency-free image metadata reader

DocxTemplater.Images now includes a built-in dependency-free image metadata reader. It reads only the metadata DocxTemplater needs for image insertion:

- image format;
- pixel width;
- pixel height;
- EXIF orientation where it is available in JPEG/TIFF metadata.

It deliberately does not decode, resize, render, transform or validate the full image payload. That keeps the adapter small and suitable for server-side document generation where DocxTemplater only needs enough information to create the correct OpenXML image part and size the drawing.

## Usage

```csharp
using DocxTemplater.Images;

template.RegisterFormatter(new ImageFormatter());
```

## Supported formats

The built-in reader supports the formats currently used by the OpenXML image service:

- JPEG
- PNG
- GIF
- BMP
- TIFF

SVG remains handled by DocxTemplater.Images directly because it can be embedded without bitmap decoding.

## Limitations

This reader is not a general image processing library. It reads headers and selected TIFF/EXIF tags only. It will not recover from unusual or corrupt image structures as comprehensively as ImageSharp, SkiaSharp or Magick.NET.

The value is that the default DocxTemplater image path does not need a third-party dependency when the required operation is metadata extraction rather than image manipulation.
