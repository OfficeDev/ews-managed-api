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

//-----------------------------------------------------------------------
// <summary>Defines the ServiceResponseCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a strogly typed list of service responses.
    /// </summary>
    /// <typeparam name="TResponse">The type of response stored in the list.</typeparam>
    [Serializable]
    public sealed class ServiceResponseCollection<TResponse> : IEnumerable<TResponse> where TResponse : ServiceResponse
    {
        private List<TResponse> responses = new List<TResponse>();
        private ServiceResult overallResult = ServiceResult.Success;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceResponseCollection&lt;TResponse&gt;"/> class.
        /// </summary>
        internal ServiceResponseCollection()
        {
        }

        /// <summary>
        /// Adds specified response.
        /// </summary>
        /// <param name="response">The response.</param>
        internal void Add(TResponse response)
        {
            EwsUtilities.Assert(
                response != null,
                "EwsResponseList.Add",
                "response is null");

            if (response.Result > this.overallResult)
            {
                this.overallResult = response.Result;
            }

            this.responses.Add(response);
        }

        /// <summary>
        /// Gets the total number of responses in the list.
        /// </summary>
        public int Count
        {
            get { return this.responses.Count; }
        }

        /// <summary>
        /// Gets the response at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the response to get.</param>
        /// <returns>The response at the specified index.</returns>
        public TResponse this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.responses[index];
            }
        }

        /// <summary>
        /// Gets a value indicating the overall result of the request that generated this response collection.
        /// If all of the responses have their Result property set to Success, OverallResult returns Success.
        /// If at least one response has its Result property set to Warning and all other responses have their Result
        /// property set to Success, OverallResult returns Warning. If at least one response has a its Result set to
        /// Error, OverallResult returns Error.
        /// </summary>
        public ServiceResult OverallResult
        {
            get { return this.overallResult; }
        }

        #region IEnumerable<TResponse>

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TResponse> GetEnumerator()
        {
            return this.responses.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (this.responses as System.Collections.IEnumerable).GetEnumerator();
        }

        #endregion
    }
}
