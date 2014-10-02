#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

namespace Microsoft.Exchange.WebServices.Data
{
    // System Dependencies
    using System.Runtime.Serialization;

    /// <summary>
    /// Source of resolution.
    /// </summary>
    public enum LocationSource
    {
        /// <summary>Unresolved</summary>
        None = 0,

        /// <summary>Resolved by external location services (such as Bing, Google, etc)</summary>
        LocationServices = 1,

        /// <summary>Resolved by external phonebook services (such as Bing, Google, etc)</summary>
        PhonebookServices = 2,

        /// <summary>Revolved by a GPS enabled device (such as cellphone)</summary>
        Device = 3,

        /// <summary>Sourced from a contact card</summary>
        Contact = 4,

        /// <summary>Sourced from a resource (such as a conference room)</summary>
        Resource = 5,
    }
}
