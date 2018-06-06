using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public sealed class FileTagModelMap : ClassMap<FileTagModel>
    {
        public FileTagModelMap()
        {
            Map(m => m.FullPath).Index(0).Default(string.Empty);
            Map(m => m.FileName).Index(1).Default(string.Empty);
            Map(m => m.Tag).Index(2).Default(string.Empty);
        }
    }

    /// <summary>
    /// Represents a row in the CSV
    /// </summary>
    public class FileTagModel
    {
        public string FullPath { get; set; }

        public string FileName { get; set; }

        public string Tag { get; set; }
    }
}
