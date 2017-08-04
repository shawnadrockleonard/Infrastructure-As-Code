using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPCustomAction
    {
        /// <summary>
        /// Defines site based custom actions
        /// </summary>
        public SPCustomActionScope Site { get; set; }

        /// <summary>
        /// Defines web based custom actions
        /// </summary>
        public SPCustomActionScope Web { get; set; }

        /// <summary>
        /// Defines list based custom actions
        /// </summary>
        public List<SPCustomActionListScope> List { get; set; }
    }

    public class SPCustomActionScope
    {
        public List<SPCustomActionBlock> scriptblocks { get; set; }

        public List<SPCustomActionLink> scriptlinks { get; set; }
    }

    public class SPCustomActionListScope
    {
        public string Title { get; set; }

        public List<SPCustomActionList> scriptcommands { get; set; }
    }

    public class SPCustomActionBase
    {
        public string name { get; set; }

        public int sequence { get; set; }
    }

    public class SPCustomActionBlock : SPCustomActionBase
    {
        public string htmlblock { get; set; }
    }

    public class SPCustomActionLink : SPCustomActionBase
    {
        public string linkurl { get; set; }
    }

    /// <summary>
    /// List Custom Action
    /// </summary>
    public class SPCustomActionList : SPCustomActionBase
    {
        /// <summary>
        /// Title of the Custom Action
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Description of the Custom Action
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Url for the Command Handler
        /// </summary>
        public string Url { get; set; }
        /// <summary>
        /// Location of the Custom Action
        /// </summary>
        public string Location { get; set; }
        /// <summary>
        /// 32px Image for the Custom Action
        /// </summary>
        public string ImageUrl { get; set; }
        /// <summary>
        /// Groupname associated by the custom action
        /// </summary>
        public string Group { get; set; }
    }

}
