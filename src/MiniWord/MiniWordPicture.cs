namespace MiniSoftware
{
    using DocumentFormat.OpenXml;

    public class MiniWordPicture
    {
        public string Path { get; set; }
        public byte[] Bytes { get; set; }
		public Int64Value Width { get; set; } = 990000L;

		public Int64Value Height { get; set; } = 792000L;
	}
}