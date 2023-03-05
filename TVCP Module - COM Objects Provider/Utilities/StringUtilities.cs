namespace TVCP_Module___COM_Objects_Provider.Utilities
{
    public static class StringUtilities
    {
        private const string NullString = "null";

        public static string NullToString(this string? source)
            => source ?? NullString;

        public static string? NullStringToNull(this string source)
            => source == NullString ? null : source;
    }
}
