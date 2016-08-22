using Slf;

namespace NODConverter
{
    public class LoggerFactory
    {
        public static ILogger GetLogger()
        {
            ILogger logger = new ConsoleLogger();
            LoggerService.SetLogger(logger);
            return LoggerService.GetLogger();
        }
    }
}
