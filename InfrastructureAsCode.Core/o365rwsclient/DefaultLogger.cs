using System;

namespace InfrastructureAsCode.Core.o365rwsclient
{
    public class DefaultLogger : ITraceLogger
    {
        public void LogError(string message)
        {
            Console.WriteLine(message);
        }

        public void LogInformation(string message)
        {
            Console.WriteLine(message);
        }
    }
}