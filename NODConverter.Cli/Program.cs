using System;
using System.Collections.Generic;
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
        private static readonly Options Options = InitOptions();

        private static readonly int[] avaliblePorts = ConfigurationHelper.GetInt32Array("AvaliblePorts");

        private const int ExitCodeConnectionFailed = 1;
        private const int ExitCodeTooFewArgs = 255;

        private static Options InitOptions()
        {
            Options options = new Options();
            options.AddOption(OptionOutputFormat);
            return options;
        }

        [STAThread]
        static void Main(string[] arguments)
        {
            ICommandLineParser commandLineParser = new PosixParser();
            CommandLine commandLine = commandLineParser.Parse(Options, arguments);

            String outputFormat = commandLine.HasOption(OptionOutputFormat.Opt) ? commandLine.GetOptionValue(OptionOutputFormat.Opt) : null;
            IDocumentFormatRegistry registry = new DefaultDocumentFormatRegistry();

            String[] fileNames = commandLine.Args;
            if ((outputFormat == null && fileNames.Length != 2) || fileNames.Length < 1)
            {
                String syntax = "Convert [options] input-file output-file; or\n" + "[options] -f output-format input-file [input-file...]";
                HelpFormatter helpFormatter = new HelpFormatter();
                helpFormatter.PrintHelp(syntax, Options);
                Environment.Exit(ExitCodeTooFewArgs);
            }

            NodProcessor nodProcessor = new NodProcessor();

            int portCount = avaliblePorts.Length,
                initialPortIndex = new Random().Next(0, portCount),
                portIndex = initialPortIndex,
                port;

            bool nodProcessed = false;

            for (int i = 0; i < portCount; i++)
            {
                port = avaliblePorts[portIndex];

                try
                {
                    nodProcessor.ProcessDocument(port, outputFormat, registry, fileNames);
                    nodProcessed = true;
                    break;
                }
                catch (ArgumentException ex)
                {
                    throw;
                }
                catch (Exception ex)
                {

                }

                portIndex = portIndex == portCount - 1 ? 0 : portIndex + 1;
            }

            if (!nodProcessed)
            {
                throw new Exception("NodProcessor unable to process Document");
            }
        }
    }
}
