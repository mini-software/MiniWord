namespace MiniSoftware.Extensions
{
    public static class StringExtension
    {
        public static string FirstCharacterToLower(this string str)
        {
            return string.IsNullOrWhiteSpace(str) ? str : char.ToLowerInvariant(str[0]) + str.Substring(1);
        }
    }
}