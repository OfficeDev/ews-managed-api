// ---------------------------------------------------------------------------
// <copyright file="UpdateInboxRulesException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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