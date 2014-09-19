// ---------------------------------------------------------------------------
// <copyright file="JsonSerializationNotImplementedException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    [Serializable]
    internal class JsonSerializationNotImplementedException : Exception
    {
        internal JsonSerializationNotImplementedException() :
            base(Strings.JsonSerializationNotImplemented)
        {
        }
    }
}
