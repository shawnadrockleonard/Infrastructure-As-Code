namespace InfrastructureAsCode.Core.o365rwsclient
{
    public interface ITraceLogger
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="message"></param>
        void LogInformation(string message);

        /// <summary>
        ///
        /// </summary>
        /// <param name="message"></param>
        void LogError(string message);
    }
}