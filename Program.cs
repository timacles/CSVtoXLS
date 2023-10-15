/* 
 - freeze top row
 - SQL to XLS
 */

class Program
{
    static void Main(string[] args)
    {
        try 
        {
            var opts = ArgParser.Parse(args);

            var parser = new CSVParser(opts.inFile);
            var converter = new XLSConverter(opts.outFile);
            
            Console.WriteLine($"Converting {opts.inFile} to {opts.outFile}");
            converter.Convert(parser.ReadLines());
            Console.WriteLine("Converting finished.");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }
}
