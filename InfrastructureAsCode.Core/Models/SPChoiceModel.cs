using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace InfrastructureAsCode.Core.Models
{
    public class SPChoiceModel
    {
        public SPChoiceModel() { }

        public SPChoiceModel(string choice, Nullable<bool> defaultChoice = null)
            : this()
        {
            this.Choice = choice;
            if (defaultChoice.HasValue)
            {
                this.DefaultChoice = defaultChoice;
            }
        }

        /// <summary>
        /// The text string
        /// </summary>
        public string Choice { get; set; }

        /// <summary>
        /// if this is the choice
        /// </summary>
        public bool? DefaultChoice { get; set; }
    }
}