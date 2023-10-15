using Microsoft.VisualBasic.FileIO;

class CSVParser 
{
    private TextFieldParser parser; 

    public CSVParser(string inFile)
    {
        parser = new TextFieldParser(inFile);
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters(",");
    }

    public IEnumerable<string[]> ReadLines()
    {
        while (!parser.EndOfData)
        {
            yield return parser.ReadFields();
        }
    }
}