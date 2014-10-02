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
// <summary>Defines the UpdateInboxRulesException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an exception thrown when an error occurs as a result of calling 
    /// the UpdateInboxRules operation.
    /// </summary>
    [Serializable]
    public sealed class UpdateInboxRulesException : ServiceRemoteException
    {
        /// <summary>
        /// ServiceResponse when service operation failed remotely.
        /// </summary>
        private ServiceResponse serviceResponse;

        /// <summary>
        /// Rule operation error collection.
        /// </summary>
        private RuleOperationErrorCollection errors;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateInboxRulesException"/> class.
        /// </summary>
        /// <param name="serviceResponse">The rule operation service response.</param>
        /// <param name="ruleOperations">The original operations.</param>
        internal UpdateInboxRulesException(UpdateInboxRulesResponse serviceResponse, IEnumerator<RuleOperation> ruleOperations)
            : base()
        {
            this.serviceResponse = serviceResponse;
            this.errors = serviceResponse.Errors;
            foreach (RuleOperationError error in this.errors)
            {
                error.SetOperationByIndex(ruleOperations);
            }
        }

        /// <summary>
        /// Gets the ServiceResponse for the exception.
        /// </summary>
        public ServiceResponse ServiceResponse
        {
            get { return this.serviceResponse; }
        }

        /// <summary>
        /// Gets the rule operation error collection.
        /// </summary>
        public RuleOperationErrorCollection Errors
        {
            get { return this.errors; }
        }

        /// <summary>
        /// Gets the rule operation error code.
        /// </summary>
        public ServiceError ErrorCode
        {
            get { return this.serviceResponse.ErrorCode; }
        }

        /// <summary>
        /// Gets the rule operation error message.
        /// </summary>
        public string ErrorMessage
        {
            get { return this.serviceResponse.ErrorMessage; }
        }
    }
}