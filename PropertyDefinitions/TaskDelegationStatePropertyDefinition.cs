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
// <summary>Defines the TaskDelegationStatePropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a task delegation property definition.
    /// </summary>
    internal sealed class TaskDelegationStatePropertyDefinition : GenericPropertyDefinition<TaskDelegationState>
    {
        private const string NoMatch = "NoMatch";
        private const string OwnNew = "OwnNew";
        private const string Owned = "Owned";
        private const string Accepted = "Accepted";

        /// <summary>
        /// Initializes a new instance of the <see cref="TaskDelegationStatePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal TaskDelegationStatePropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
        }

        /// <summary>
        /// Parses the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>TaskDelegationState value.</returns>
        internal override object Parse(string value)
        {
            switch (value)
            {
                case NoMatch:
                    return TaskDelegationState.NoDelegation;
                case OwnNew:
                    return TaskDelegationState.Unknown;
                case Owned:
                    return TaskDelegationState.Accepted;
                case Accepted:
                    return TaskDelegationState.Declined;
                default:
                    EwsUtilities.Assert(
                        false,
                        "TaskDelegationStatePropertyDefinition.Parse",
                        string.Format("TaskDelegationStatePropertyDefinition.Parse(): value {0} cannot be handled.", value));
                    return null; // To keep the compiler happy
            }
        }

        /// <summary>
        /// Convert instance to string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>TaskDelegationState value.</returns>
        internal override string ToString(object value)
        {
            TaskDelegationState taskDelegationState = (TaskDelegationState)value;

            switch (taskDelegationState)
            {
                case TaskDelegationState.NoDelegation:
                    return NoMatch;
                case TaskDelegationState.Unknown:
                    return OwnNew;
                case TaskDelegationState.Accepted:
                    return Owned;
                case TaskDelegationState.Declined:
                    return Accepted;
                default:
                    EwsUtilities.Assert(
                        false,
                        "TaskDelegationStatePropertyDefinition.ToString",
                        "Invalid TaskDelegationState value.");
                    return null; // To keep the compiler happy
            }
        }
    }
}
