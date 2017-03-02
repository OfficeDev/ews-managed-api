using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// The action to perform when the item is uploaded.
    /// </summary>
    public enum CreateAction
    {
        /// <summary>
        /// Create a new item
        /// </summary>
        CreateNew,

        /// <summary>
        /// Update the item if it already exists
        /// </summary>

        Update
    }
}
