using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.Enums
{
    public enum ViewToolBarEnum
    {
        /// <summary>
        /// The most common type of toolbar, which is used, for example, in the All Items views for most lists, and which corresponds to Full Toolbar in the Web Part tool pane.
        /// </summary>
        Standard,

        /// <summary>
        /// Used in Default.aspx and Web Part Pages and corresponds to Summary Toolbar in the Web Part tool pane.
        /// </summary>
        FreeForm,

        /// <summary>
        /// No toolbar is used in the view, corresponding to No Toolbar in the Web Part tool pane.
        /// </summary>
        None
    }
}
