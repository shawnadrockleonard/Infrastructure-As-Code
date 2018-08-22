using System;

namespace InfrastructureAsCode.Powershell.Commands.Base
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class CmdletHelpAttribute : Attribute
    {
        readonly string description;

        public string DetailedDescription { get; set; }

        public string Copyright { get; set; }

        public string Version { get; set; }

        public string Category { get; set; }

        public CmdletHelpAttribute(string description)
        {
            this.description = description;
        }

        public string Description
        {
            get
            {
                return description;
            }
        }
    }

}