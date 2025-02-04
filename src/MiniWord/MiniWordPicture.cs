using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MiniSoftware.Common.Enums;
using MiniSoftware.Extensions;

namespace MiniSoftware
{
    public enum MiniWordPictureWrappingType
    {
        Inline,
        Anchor
    }

    public class MiniWordPicture
    {
        public MiniWordPicture(string path, long width = 400, long height = 400)
        {
            Path = path;
            Width = width;
            Height= height;
        }

        public MiniWordPicture(byte[] bytes, Extension extension, long width = 400, long height = 400)
        {
            Bytes = bytes;
            Extension = extension.ToString().FirstCharacterToLower();

            Width = width;
            Height = height;
        }

        private string _extension;

        public MiniWordPictureWrappingType WrappingType { get; set; } = MiniWordPictureWrappingType.Inline;

        public bool BehindDoc { get; set; } = false;
        public bool AllowOverlap { get; set; } = false;
        public long HorizontalPositionOffset { get; set; } = 0;
        public long VerticalPositionOffset { get; set; } = 0;

        public string Path { get; set; }

        public string Extension
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_extension))
                    return _extension;

                if (Path != null)
                    return System.IO.Path.GetExtension(Path).ToUpperInvariant().Replace(".", "");

                return string.Empty;
            }
            set => _extension = value;
        }

        internal PartTypeInfo GetImagePartType
        {
            get
            {
                switch (Extension.ToLower())
                {
                    case "bmp": return ImagePartType.Bmp;
                    case "emf": return ImagePartType.Emf;
                    case "ico": return ImagePartType.Icon;
                    case "jpg":
                    case "jpeg":
                        return ImagePartType.Jpeg;
                    case "pcx": return ImagePartType.Pcx;
                    case "png": return ImagePartType.Png;
                    case "svg": return ImagePartType.Svg;
                    case "tiff": return ImagePartType.Tiff;
                    case "wmf": return ImagePartType.Wmf;
                    default:
                        throw new NotSupportedException($"{_extension} is not supported");
                }
            }
        }

        public byte[] Bytes { get; set; }

        /// <summary>
        ///     Unit is Pixel
        /// </summary>
        public Int64Value Width { get; set; }
        
        internal Int64Value Cx => Width * 9525;

        /// <summary>
        ///     Unit is Pixel
        /// </summary>
        public Int64Value Height { get; set; }

        //format resource from http://openxmltrix.blogspot.com/2011/04/updating-images-in-image-placeholde-and.html
        internal Int64Value Cy => Height * 9525;
    }
}