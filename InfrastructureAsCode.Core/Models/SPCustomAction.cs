using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPCustomAction
    {
        public SPCustomActionScope Site { get; set; }

        public SPCustomActionScope Web { get; set; }
    }

    public class SPCustomActionScope
    {
        public List<SPCustomActionBlock> scriptblocks { get; set; }

        public List<SPCustomActionLink> scriptlinks { get; set; }
    }

    public class SPCustomActionBlock : SPCustomActionBase
    {
        public string htmlblock { get; set; }
    }

    public class SPCustomActionLink : SPCustomActionBase
    {
        public string linkurl { get; set; }
    }

    public class SPCustomActionBase
    {
        public string name { get; set; }

        public int sequence { get; set; }
    }
}
