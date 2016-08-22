using System;
using System.IO;
using net.sf.dotnetcli;
using NODConverter.OpenOffice.Connection;
using NODConverter.OpenOffice.Converter;

namespace NODConverter.Cli
{
    /// <summary>
    /// Command line tool to Convert a document into a different format.
    /// <p/>
    /// Usage: can Convert a single file
    /// 
    /// <pre>
    /// ConvertDocument test.odt test.pdf
    /// </pre>
    /// 
    /// or multiple files at once by specifying the output format
    /// 
    /// <pre>
    /// ConvertDocument -f pdf test1.odt test2.odt
    /// ConvertDocument -f pdf *.odt
    /// </pre>
    /// </summary>
    class Program
    {
        private static readonly Option OptionOutputFormat = new Option("f", "output-format", true, "output format (e.g. pdf)");
        private static readonly Option OptionPort = new Option("p", "port", true, "OpenOffice.org port");
        private static readonly Option OptionVerbose = new Option("v", "verbose", false, "verbose");
        private static readonly Option OptionXmlRegistry = new Option("x", "xml-registry", true, "XML document format registry");
        private static readonly Options Options = InitOptions();

        private const int ExitCodeConnectionFailed = 1;
        //private const int ExitCodeXmlRegistryNotFound = 254;
        private const int ExitCodeTooFewArgs = 255;

        private static Options InitOptions()
        {
            Options options = new Options();
            options.AddOption(OptionOutputFormat);
            options.AddOption(OptionPort);
            options.AddOption(OptionVerbose);
            options.AddOption(OptionXmlRegistry);
            return options;
        }
        [STAThread]
        static void Main(string[] arguments)
        {
            ICommandLineParser commandLineParser = new PosixParser();
            CommandLine commandLine = commandLineParser.Parse(Options, arguments);

            int port = SocketOpenOfficeConnection.DefaultPort;
            if (commandLine.HasOption(OptionPort.Opt))
            {
                port = Convert.ToInt32(commandLine.GetOptionValue(OptionPort.Opt));
            }

            String outputFormat = null;
            if (commandLine.HasOption(OptionOutputFormat.Opt))
            {
                outputFormat = commandLine.GetOptionValue(OptionOutputFormat.Opt);
            }

            bool verbose = commandLine.HasOption(OptionVerbose.Opt);
  
            IDocumentFormatRegistry registry = new DefaultDocumentFormatRegistry();

            String[] fileNames = commandLine.Args;
            if ((outputFormat == null && fileNames.Length != 2) || fileNames.Length < 1)
            {
                String syntax = "Convert [options] input-file output-file; or\n" + "[options] -f output-format input-file [input-file...]";
                HelpFormatter helpFormatter = new HelpFormatter();
                helpFormatter.PrintHelp(syntax, Options);
                Environment.Exit(ExitCodeTooFewArgs);
            }

            IOpenOfficeConnection connection = new SocketOpenOfficeConnection(port);
            try
            {
                if (verbose)
                {
                    Console.Out.WriteLine("-- connecting to OpenOffice.org on port " + port);
                }
                connection.Connect();
            }
            catch (Exception)
            {
                Console.Error.WriteLine("ERROR: connection failed. Please make sure OpenOffice.org is running and listening on port " + port + ".");
                Environment.Exit(ExitCodeConnectionFailed);
            }
            try
            {
                IDocumentConverter converter = new OpenOfficeDocumentConverter(connection, registry);
                if (outputFormat == null)
                {
                    FileInfo inputFile = new FileInfo(fileNames[0]);
                    FileInfo outputFile = new FileInfo(fileNames[1]);
                    ConvertOne(converter, inputFile, outputFile, verbose);
                }
                else
                {
                    foreach (var t in fileNames)
                    {
                        var inputFile = new FileInfo(t);
                        var outputFile = new FileInfo(inputFile.FullName.Remove(inputFile.FullName.LastIndexOf(".", StringComparison.Ordinal)) + "." + outputFormat);
                        ConvertOne(converter, inputFile, outputFile, verbose);
                    }
                }
            }
            finally
            {
                if (verbose)
                {
                    Console.Out.WriteLine("-- disconnecting");
                }
                connection.Disconnect();
            }
        }
        private static void ConvertOne(IDocumentConverter converter, FileInfo inputFile, FileInfo outputFile, bool verbose)
        {
            if (verbose)
            {
                Console.Out.WriteLine("-- converting " + inputFile + " to " + outputFile);
            }
            converter.Convert(inputFile, outputFile);
        }
    }
}
