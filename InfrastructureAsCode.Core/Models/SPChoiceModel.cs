using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace InfrastructureAsCode.Core.Models
{
    public class SPChoiceModel
    {
        public SPChoiceModel() { }

        public SPChoiceModel(string choice, bool defaultChoice = false)
            : this()
        {
            this.Choice = choice;
            this.DefaultChoice = defaultChoice;
        }

        /// <summary>
        /// The text string
        /// </summary>
        public string Choice { get; set; }

        /// <summary>
        /// if this is the choice
        /// </summary>
        public bool DefaultChoice { get; set; }
    }
}