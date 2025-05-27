using DocumentFormat.OpenXml;

namespace DocxTemplater.ImageBase
{
    public sealed class ImageRotation
    {
        private const int MUnitsToDegree = 60000; // 1/60000 degree

        private ImageRotation(int units)
        {
            Units = units;
        }

        public int Units { get; private set; }

        public static ImageRotation CreateFromDegree(int degree)
        {
            int units = Mod(degree, 360) * MUnitsToDegree;
            return new ImageRotation(units);
        }

        public static ImageRotation CreateFromUnits(int units)
        {
            return new ImageRotation(units);
        }

        /// <summary>
        /// -90 degree is 270 degree - not normal C# % operator because it returns negative value
        /// </summary>
        private static int Mod(int a, int n)
        {
            return ((a % n) + n) % n;
        }

        public static ImageRotation CreateFromExifRotation(ExifRotation orientationValueValue)
        {
            return orientationValueValue switch
            {
                ExifRotation.Rotate90 => CreateFromDegree(90),
                ExifRotation.Rotate180 => CreateFromDegree(180),
                ExifRotation.Rotate270 => CreateFromDegree(270),
                _ => CreateFromDegree(0)
            };
        }

        public ImageRotation AddUnits(Int32Value transformRotation)
        {
            if (transformRotation == null)
            {
                return this;
            }

            var rotation = transformRotation.Value;
            if (rotation == 0)
            {
                return this;
            }

            var newUnits = Units + rotation;
            //units is in 1/60000 degree  - modulo
            newUnits = (int)Mod(newUnits, 360 * MUnitsToDegree);
            return new ImageRotation(newUnits);
        }
    }


    public enum ExifRotation
    {
        Normal = 1,
        Rotate180 = 3,
        Rotate90 = 6,
        Rotate270 = 8
    }
}
