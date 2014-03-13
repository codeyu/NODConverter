using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter
{
    public class LoggerFactory
    {
        public static Slf.ILogger GetLogger()
        {
            Slf.ILogger logger = new Slf.ConsoleLogger();
            Slf.LoggerService.SetLogger(logger);
            return Slf.LoggerService.GetLogger();
        }
    }
}
