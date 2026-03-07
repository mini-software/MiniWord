using DocumentFormat.OpenXml.Wordprocessing;

namespace MiniSoftware
{
    /// <summary>
    /// Represents a single border edge (top, bottom, left, right, inside horizontal, inside vertical).
    /// </summary>
    public class MiniWordBorder
    {
        /// <summary>Border line style. Defaults to <see cref="BorderValues.Single"/>.</summary>
        public BorderValues Style { get; set; } = BorderValues.Single;

        /// <summary>Border color as a hex string, e.g. "000000" or "#000000".</summary>
        public string Color { get; set; } = "000000";

        /// <summary>Border width in eighths of a point. Defaults to 4 (= 0.5 pt).</summary>
        public uint Size { get; set; } = 4;

        /// <summary>Space between the border and text, in points. Defaults to 0.</summary>
        public uint Space { get; set; } = 0;
    }

    /// <summary>
    /// Specifies table-level border settings that MiniWord applies when the tag value is
    /// a <see cref="MiniWordTableBorder"/> instance.  Place <c>{{tagName}}</c> anywhere
    /// inside a table cell; MiniWord will apply the borders to the parent table and clear
    /// the placeholder text.
    /// </summary>
    public class MiniWordTableBorder
    {
        /// <summary>Top outer border of the table.</summary>
        public MiniWordBorder TopBorder { get; set; }

        /// <summary>Bottom outer border of the table.</summary>
        public MiniWordBorder BottomBorder { get; set; }

        /// <summary>Left outer border of the table.</summary>
        public MiniWordBorder LeftBorder { get; set; }

        /// <summary>Right outer border of the table.</summary>
        public MiniWordBorder RightBorder { get; set; }

        /// <summary>Horizontal border between rows (inside the table).</summary>
        public MiniWordBorder InsideHorizontalBorder { get; set; }

        /// <summary>Vertical border between columns (inside the table).</summary>
        public MiniWordBorder InsideVerticalBorder { get; set; }

        /// <summary>
        /// Returns a <see cref="MiniWordTableBorder"/> where all six border edges use the
        /// same style, color and size.
        /// </summary>
        /// <param name="color">Hex color string, e.g. "#000000" or "000000".</param>
        /// <param name="size">Border width in eighths of a point (default 4 = 0.5 pt).</param>
        /// <param name="style">Border line style (default <see cref="BorderValues.Single"/>).</param>
        public static MiniWordTableBorder Create(string color = "000000", uint size = 4)
        {
            return Create(color, size, BorderValues.Single);
        }

        /// <summary>
        /// Returns a <see cref="MiniWordTableBorder"/> where all six border edges use the
        /// same style, color and size.
        /// </summary>
        /// <param name="color">Hex color string, e.g. "#000000" or "000000".</param>
        /// <param name="size">Border width in eighths of a point.</param>
        /// <param name="style">Border line style.</param>
        public static MiniWordTableBorder Create(string color, uint size, BorderValues style)
        {
            var normalizedColor = color?.Replace("#", "") ?? "000000";
            return new MiniWordTableBorder
            {
                TopBorder              = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style },
                BottomBorder           = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style },
                LeftBorder             = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style },
                RightBorder            = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style },
                InsideHorizontalBorder = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style },
                InsideVerticalBorder   = new MiniWordBorder { Color = normalizedColor, Size = size, Style = style }
            };
        }
    }
}
