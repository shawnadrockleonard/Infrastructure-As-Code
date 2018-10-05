namespace InfrastructureAsCode.Core.Models
{
    public class CollectionModel
    {
        /// <summary>
        /// Absolute URL
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Total subweb count
        /// </summary>
        public int WebsCount { get; set; }

        /// <summary>
        /// emit URLs
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Url;
        }
    }
}