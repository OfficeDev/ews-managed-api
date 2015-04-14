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
    using System.Collections.Generic;
    using System.Xml;
    
    /// <summary>
    /// Represents the InsightValue.
    /// </summary>
    public class InsightValue : ComplexProperty
    {
        private InsightSourceType insightSource;
        private long updatedUtcTicks;

        /// <summary>
        /// Gets the InsightSource
        /// </summary>
        public InsightSourceType InsightSource
        {
            get
            {
                return this.insightSource;
            }

            set
            {
                this.SetFieldValue<InsightSourceType>(ref this.insightSource, value);
            }
        }

        /// <summary>
        /// Gets the UpdatedUtcTicks
        /// </summary>
        public long UpdatedUtcTicks
        {
            get
            {
                return this.updatedUtcTicks;
            }

            set
            {
                this.SetFieldValue<long>(ref this.updatedUtcTicks, value);
            }
        }
    }
}