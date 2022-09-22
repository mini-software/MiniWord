namespace MiniSoftware
{
    using DocumentFormat.OpenXml.Wordprocessing;

    public class MiniWorHyperLink{
        public string Url { get; set; }

        public string Text { get; set; }

        public UnderlineValues UnderLineValue { get; set; } = UnderlineValues.Single;

        public TargetFrameType TargetFrame { get; set; } = TargetFrameType.Blank;

        internal string GetTargetFrame(){
            
            switch (TargetFrame)
            {
                case TargetFrameType.Blank:
                    return "_blank";
                case TargetFrameType.Top:
                    return "_top";
                case TargetFrameType.Self:
                    return "_self";
                case TargetFrameType.Parent:
                    return "_parent";
            }

            return "_blank";
        }
    }
}