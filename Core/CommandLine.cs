using CommandLine;
using CommandLine.Text;

namespace DocxToPdf.Core
{

    /// <summary>
    /// This class allow easy parsing of command line arguments.
    /// </summary>
    public class CommandLineArguments
    {
        [Option('i', "input", Required = true, HelpText = "Input .docx file which will be converted.")]
        public string InputFile { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output .pdf file.")]
        public string OutputFile { get; set; }

        [Option('s', "silent", DefaultValue = false, HelpText = "Remove all log messages from output.")]
        public bool Silent { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
