namespace MiniSoftware
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using System;
    using System.Collections.Generic;

    public enum MiniWordPictureWrappingType
    {
        Inline,
        Anchor
    }
    
    public class MiniWordPicture
    {
        
        public MiniWordPictureWrappingType WrappingType { get; set; } = MiniWordPictureWrappingType.Inline;

        public bool BehindDoc { get; set; } = false;
        public bool AllowOverlap { get; set; } = false;
        public long HorizontalPositionOffset { get; set; } = 0;
        public long VerticalPositionOffset { get; set; } = 0;
        
        public string Path { get; set; }
        private string _extension;
        public string Extension
        {
            get
            {
                if (Path != null)
                    return System.IO.Path.GetExtension(Path).ToUpperInvariant().Replace(".", "");
                else
                {
                    return _extension.ToUpper();
                }
            }
            set { _extension = value; }
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
                    case "jpg": return ImagePartType.Jpeg;
                    case "jpeg": return ImagePartType.Jpeg;
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
        /// Unit is Pixel
        /// </summary>
		public Int64Value Width { get; set; } = 400;
        internal Int64Value Cx { get { return Width * 9525; } }
        /// <summary>
        /// Unit is Pixel
        /// </summary>
        public Int64Value Height { get; set; } = 400;
        //format resource from http://openxmltrix.blogspot.com/2011/04/updating-images-in-image-placeholde-and.html
        internal Int64Value Cy { get { return Height * 9525; } }
    }
}