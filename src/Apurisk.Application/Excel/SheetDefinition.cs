namespace Apurisk.Application.Excel
{
    public sealed class SheetDefinition
    {
        public SheetDefinition(string name, string[] headers)
        {
            Name = name;
            Headers = headers;
        }

        public string Name { get; private set; }
        public string[] Headers { get; private set; }
    }
}
