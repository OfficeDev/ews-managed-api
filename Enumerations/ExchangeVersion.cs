/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Defines the each available Exchange release version
    /// </summary>
    public enum ExchangeVersion
    {
        /// <summary>
        /// Microsoft Exchange 2007, Service Pack 1
        /// </summary>
        Exchange2007_SP1 = 0,

        /// <summary>
        /// Microsoft Exchange 2010
        /// </summary>
        Exchange2010 = 1,

        /// <summary>
        /// Microsoft Exchange 2010, Service Pack 1
        /// </summary>
        Exchange2010_SP1 = 2,

        /// <summary>
        /// Microsoft Exchange 2010, Service Pack 2
        /// </summary>
        Exchange2010_SP2 = 3,

        /// <summary>
        /// Microsoft Exchange 2013
        /// </summary>
        Exchange2013 = 4,

        /// <summary>
        /// Microsoft Exchange 2013 SP1
        /// </summary>
        Exchange2013_SP1 = 5,

        /// <summary>
        /// Microsoft Exchange 2016
        /// </summary>
        Exchange2016 = 6,
    }
}