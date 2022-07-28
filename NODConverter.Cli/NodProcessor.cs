using NODConverter.OpenOffice.Connection;
using NODConverter.OpenOffice.Converter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NODConverter.Cli
{
    internal class NodProcessor
    {
        private const int ExitCodeConnectionFailed = 1;
        private const int ExitCodeTooFewArgs = 255;

        public void ProcessDocument(int port, string outputFormat, IDocumentFormatRegistry registry, string[] fileNames) {

            IOpenOfficeConnection connection = new SocketOpenOfficeConnection(port);
            OfficeInfo oo = EnvUtils.Get();
            if (oo.Kind == OfficeKind.Unknown)
            {
                Console.Out.WriteLine("please setup OpenOffice or LibreOffice!");
                return;
            }
            try
            {
                Console.Out.WriteLine("connecting to OpenOffice.org on port " + port);

                // Connect to existing instance

                connection.Connect();

            }
            catch (Exception ex)
            {

                // Cannot connect to existing instance - start a new one

                string CmdArguments = string.Format("-headless -accept=\"socket,host={0},port={1};urp;\" -nofirststartwizard", SocketOpenOfficeConnection.DefaultHost, port);
                if (!EnvUtils.RunCmd(oo.OfficeUnoPath, "soffice", CmdArguments))
                {
                    Console.Error.WriteLine("ERROR: connection failed. Please make sure OpenOffice.org is running and listening on port " + port + ".");
                    Environment.Exit(ExitCodeConnectionFailed);
                }

            }
            try
            {

                IDocumentConverter converter = new OpenOfficeDocumentConverter(connection, registry);
                if (outputFormat == null)
                {
                    FileInfo inputFile = new FileInfo(fileNames[0]);
                    FileInfo outputFile = new FileInfo(fileNames[1]);
                    ConvertOne(converter, inputFile, outputFile);
                }
                else
                {
                    foreach (var t in fileNames)
                    {
                        var inputFile = new FileInfo(t);
                        var outputFile = new FileInfo(inputFile.FullName.Remove(inputFile.FullName.LastIndexOf(".", StringComparison.Ordinal)) + "." + outputFormat);
                        ConvertOne(converter, inputFile, outputFile);
                    }
                }
            }
            finally
            {
                Console.Out.WriteLine("disconnecting from OpenOffice.org on port " + port);
                connection.Disconnect();
            }
        }

        private static void ConvertOne(IDocumentConverter converter, FileInfo inputFile, FileInfo outputFile)
        {
            Console.Out.WriteLine("converting " + inputFile + " to " + outputFile);
            converter.Convert(inputFile, outputFile);
        }
    }
}
