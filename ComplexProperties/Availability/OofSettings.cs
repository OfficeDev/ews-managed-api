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
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a user's Out of Office (OOF) settings.
    /// </summary>
    public sealed class OofSettings : ComplexProperty, ISelfValidate
    {
        private OofState state;
        private OofExternalAudience externalAudience;
        private OofExternalAudience allowExternalOof;
        private TimeWindow duration;
        private OofReply internalReply;
        private OofReply externalReply;

        /// <summary>
        /// Serializes an OofReply. Emits an empty OofReply in case the one passed in is null.
        /// </summary>
        /// <param name="oofReply">The oof reply.</param>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        private void SerializeOofReply(
            OofReply oofReply,
            EwsServiceXmlWriter writer,
            string xmlElementName)
        {
            if (oofReply != null)
            {
                oofReply.WriteToXml(writer, xmlElementName);
            }
            else
            {
                OofReply.WriteEmptyReplyToXml(writer, xmlElementName);
            }
        }

        /// <summary>
        /// Initializes a new instance of OofSettings.
        /// </summary>
        public OofSettings()
            : base()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.OofState:
                    this.state = reader.ReadValue<OofState>();
                    return true;
                case XmlElementNames.ExternalAudience:
                    this.externalAudience = reader.ReadValue<OofExternalAudience>();
                    return true;
                case XmlElementNames.Duration:
                    this.duration = new TimeWindow();
                    this.duration.LoadFromXml(reader);
                    return true;
                case XmlElementNames.InternalReply:
                    this.internalReply = new OofReply();
                    this.internalReply.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ExternalReply:
                    this.externalReply = new OofReply();
                    this.externalReply.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.OofState,
                this.State);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.ExternalAudience,
                this.ExternalAudience);

            if (this.Duration != null && this.State == OofState.Scheduled)
            {
                this.Duration.WriteToXml(writer, XmlElementNames.Duration);
            }

            this.SerializeOofReply(
                this.InternalReply,
                writer,
                XmlElementNames.InternalReply);
            this.SerializeOofReply(
                this.ExternalReply,
                writer,
                XmlElementNames.ExternalReply);
        }

        /// <summary>
        /// Gets or sets the user's OOF state.
        /// </summary>
        /// <value>The user's OOF state.</value>
        public OofState State
        {
            get { return this.state; }
            set { this.state = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating who should receive external OOF messages.
        /// </summary>
        public OofExternalAudience ExternalAudience
        {
            get { return this.externalAudience; }
            set { this.externalAudience = value; }
        }

        /// <summary>
        /// Gets or sets the duration of the OOF status when State is set to OofState.Scheduled.
        /// </summary>
        public TimeWindow Duration
        {
            get { return this.duration; }
            set { this.duration = value; }
        }

        /// <summary>
        /// Gets or sets the OOF response sent other users in the user's domain or trusted domain.
        /// </summary>
        public OofReply InternalReply
        {
            get { return this.internalReply; }
            set { this.internalReply = value; }
        }

        /// <summary>
        /// Gets or sets the OOF response sent to addresses outside the user's domain or trusted domain.
        /// </summary>
        public OofReply ExternalReply
        {
            get { return this.externalReply; }
            set { this.externalReply = value; }
        }

        /// <summary>
        /// Gets a value indicating the authorized external OOF notifications.
        /// </summary>
        public OofExternalAudience AllowExternalOof
        {
            get { return this.allowExternalOof; }
            internal set { this.allowExternalOof = value; }
        }

        #region ISelfValidate Members

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            if (this.State == OofState.Scheduled)
            {
                if (this.Duration == null)
                {
                    throw new ArgumentException(Strings.DurationMustBeSpecifiedWhenScheduled);
                }

                EwsUtilities.ValidateParam(this.Duration, "Duration");
            }
        }

        #endregion
    }
}