using System;
using System.Collections.Generic;
using System.Text;
using net.sf.dotnetcli;
using NODConverter.OpenOffice.Connection;
using NODConverter.OpenOffice.Converter;
namespace NODConverter.Cli
{
    /// <summary>
    /// Command line tool to convert a document into a different format.
    /// <p>
    /// Usage: can convert a single file
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
        private static readonly Option OPTION_OUTPUT_FORMAT = new Option("f", "output-format", true, "output format (e.g. pdf)");
        private static readonly Option OPTION_PORT = new Option("p", "port", true, "OpenOffice.org port");
        private static readonly Option OPTION_VERBOSE = new Option("v", "verbose", false, "verbose");
        private static readonly Option OPTION_XML_REGISTRY = new Option("x", "xml-registry", true, "XML document format registry");
        private static readonly Options OPTIONS = initOptions();

        private const int EXIT_CODE_CONNECTION_FAILED = 1;
        private const int EXIT_CODE_XML_REGISTRY_NOT_FOUND = 254;
        private const int EXIT_CODE_TOO_FEW_ARGS = 255;

        private static Options initOptions()
        {
            Options options = new Options();
            options.AddOption(OPTION_OUTPUT_FORMAT);
            options.AddOption(OPTION_PORT);
            options.AddOption(OPTION_VERBOSE);
            options.AddOption(OPTION_XML_REGISTRY);
            return options;
        }
        [STAThread]
        static void Main(string[] arguments)
        {
            ICommandLineParser commandLineParser = new PosixParser();
            CommandLine commandLine = commandLineParser.Parse(OPTIONS, arguments);

            int port = SocketOpenOfficeConnection.DEFAULT_PORT;
            if (commandLine.HasOption(OPTION_PORT.Opt))
            {
                port = Convert.ToInt32(commandLine.GetOptionValue(OPTION_PORT.Opt));
            }

            System.String outputFormat = null;
            if (commandLine.HasOption(OPTION_OUTPUT_FORMAT.Opt))
            {
                outputFormat = commandLine.GetOptionValue(OPTION_OUTPUT_FORMAT.Opt);
            }

            bool verbose = false;
            if (commandLine.HasOption(OPTION_VERBOSE.Opt))
            {
                verbose = true;
            }

            IDocumentFormatRegistry registry;
            //if (commandLine.HasOption(OPTION_XML_REGISTRY.Opt))
            //{
            //    System.IO.FileInfo registryFile = new FileInfo(commandLine.GetOptionValue(OPTION_XML_REGISTRY.Opt));
            //    if (!System.IO.File.Exists(registryFile.FullName))
            //    {
            //        System.Console.Error.WriteLine("ERROR: specified XML registry file not found: " + registryFile);
            //        System.Environment.Exit(EXIT_CODE_XML_REGISTRY_NOT_FOUND);
            //    }
            //    //UPGRADE_TODO: 构造函数“java.io.FileInputStream.FileInputStream”被转换为具有不同行为的 'System.IO.FileStream.FileStream'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javaioFileInputStreamFileInputStream_javaioFile'"
            //    //registry = new XmlDocumentFormatRegistry(new System.IO.FileStream(registryFile.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read));
            //    if (verbose)
            //    {
            //        System.Console.Out.WriteLine("-- loaded document format registry from " + registryFile);
            //    }
            //}
            //else
            //{
            registry = new DefaultDocumentFormatRegistry();
            //}

            System.String[] fileNames = commandLine.Args;
            if ((outputFormat == null && fileNames.Length != 2) || fileNames.Length < 1)
            {
                System.String syntax = "convert [options] input-file output-file; or\n" + "[options] -f output-format input-file [input-file...]";
                HelpFormatter helpFormatter = new HelpFormatter();
                helpFormatter.PrintHelp(syntax, OPTIONS);
                System.Environment.Exit(EXIT_CODE_TOO_FEW_ARGS);
            }

            IOpenOfficeConnection connection = new SocketOpenOfficeConnection(port);
            try
            {
                if (verbose)
                {
                    System.Console.Out.WriteLine("-- connecting to OpenOffice.org on port " + port);
                }
                connection.connect();
            }
            catch (System.Exception officeNotRunning)
            {
                System.Console.Error.WriteLine("ERROR: connection failed. Please make sure OpenOffice.org is running and listening on port " + port + ".");
                System.Environment.Exit(EXIT_CODE_CONNECTION_FAILED);
            }
            try
            {
                IDocumentConverter converter = new OpenOfficeDocumentConverter(connection, registry);
                if (outputFormat == null)
                {
                    System.IO.FileInfo inputFile = new System.IO.FileInfo(fileNames[0]);
                    System.IO.FileInfo outputFile = new System.IO.FileInfo(fileNames[1]);
                    convertOne(converter, inputFile, outputFile, verbose);
                }
                else
                {
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        System.IO.FileInfo inputFile = new System.IO.FileInfo(fileNames[i]);
                        //System.IO.FileInfo outputFile = new System.IO.FileInfo((FilenameUtils.getFullPath(fileNames[i]) + FilenameUtils.getBaseName(fileNames[i])) + "." + outputFormat);
                        System.IO.FileInfo outputFile = new System.IO.FileInfo(inputFile.FullName.Remove(inputFile.FullName.LastIndexOf(".")) + "." + outputFormat);
                        convertOne(converter, inputFile, outputFile, verbose);
                    }
                }
            }
            finally
            {
                if (verbose)
                {
                    System.Console.Out.WriteLine("-- disconnecting");
                }
                connection.disconnect();
            }
        }
        private static void convertOne(IDocumentConverter converter, System.IO.FileInfo inputFile, System.IO.FileInfo outputFile, bool verbose)
        {
            if (verbose)
            {
                System.Console.Out.WriteLine("-- converting " + inputFile + " to " + outputFile);
            }
            converter.convert(inputFile, outputFile);
        }
    }
}
