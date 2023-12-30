namespace SmallWords.Replacing
{
    public class WordsReplacer
    {

        public static void Replace(Stream stream, IDataReplacer replacer, Dictionary<string, object> data)
        {
            replacer.Replace(stream, data);
        }

        public static void Replace(string documentPath, IDataReplacer replacer, Dictionary<string, object> data)
        {
            using (var fs = File.Open(documentPath, FileMode.Open, FileAccess.ReadWrite))
            {
                replacer.Replace(fs, data);
            }
        }

    }
}
