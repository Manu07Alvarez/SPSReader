namespace SPSReader.types;

public  class XlSheets
{

        public List<Dictionary<string, List<Dictionary<string, string>>>> Sheets { get; set; } = [];

        public Dictionary<string, List<Dictionary<string, string>>> Sheet { get; set; } = new();

}