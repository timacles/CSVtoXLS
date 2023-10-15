using CommandLine;

public class Options
{
    [Option('i', "inFile", Required = true, HelpText = "CSV File to convert")]
    public string? inFile { get; set; }

    [Option('o', "outFile", Required = false, HelpText = "Filename for output file")]
    public string? outFile { get; set; }
}

public class ArgParser
{
    public static Options Parse(string[] args)
    {
        Options opts = Parser.Default.ParseArguments<Options>(args).Value;
        Validate(opts);
        return opts;
    }

    private static void Validate(Options opts)
    {
        if (!File.Exists(opts?.inFile))
           throw new ArgParseException($"Input File {opts.inFile} does not exist.");

        if (!opts.inFile.EndsWith(".csv"))
           throw new ArgParseException($"Not a CSV file: {opts.inFile}");

        if (opts.outFile is null)
            opts.outFile = opts.inFile.Replace(".csv", ".xlsx");

        if (File.Exists(opts.outFile))
           throw new ArgParseException($"Output File {opts.outFile} already exist.");
    }
}

public class ArgParseException : Exception
{
    public ArgParseException(string message) : base(message) { }
}