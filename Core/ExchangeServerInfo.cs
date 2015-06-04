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
    /// Represents Exchange server information.
    /// </summary>
    public sealed class ExchangeServerInfo
    {
        private int majorVersion;
        private int minorVersion;
        private int majorBuildNumber;
        private int minorBuildNumber;
        private string versionString;
        
        /// <summary>
        /// Default constructor
        /// </summary>
        internal ExchangeServerInfo()
        {
        }

        /// <summary>
        /// Parse current element to extract server information
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        /// <returns>ExchangeServerInfo</returns>
        internal static ExchangeServerInfo Parse(EwsServiceXmlReader reader)
        {
            EwsUtilities.Assert(
                                reader.HasAttributes, 
                                "ExchangeServerVersion.Parse",
                                "Current element doesn't have attributes");

            ExchangeServerInfo info = new ExchangeServerInfo();
            info.MajorVersion = reader.ReadAttributeValue<int>("MajorVersion");
            info.MinorVersion = reader.ReadAttributeValue<int>("MinorVersion");
            info.MajorBuildNumber = reader.ReadAttributeValue<int>("MajorBuildNumber");
            info.MinorBuildNumber = reader.ReadAttributeValue<int>("MinorBuildNumber");
            info.VersionString = reader.ReadAttributeValue("Version");
            return info;
        }

        /// <summary>
        /// Gets the Major Exchange server version number
        /// </summary>
        public int MajorVersion
        {
            get { return this.majorVersion; }
            internal set { this.majorVersion = value; }
        }

        /// <summary>
        /// Gets the Minor Exchange server version number
        /// </summary>
        public int MinorVersion
        {
            get { return this.minorVersion; }
            internal set { this.minorVersion = value; }
        }

        /// <summary>
        /// Gets the Major Exchange server build number
        /// </summary>
        public int MajorBuildNumber
        {
            get { return this.majorBuildNumber; }
            internal set { this.majorBuildNumber = value; }
        }

        /// <summary>
        /// Gets the Minor Exchange server build number
        /// </summary>
        public int MinorBuildNumber
        {
            get { return this.minorBuildNumber; }
            internal set { this.minorBuildNumber = value; }
        }

        /// <summary>
        /// Gets the Exchange server version string (e.g. "Exchange2010")
        /// </summary>
        /// <remarks>
        /// The version is a string rather than an enum since its possible for the client to
        /// be connected to a later server for which there would be no appropriate enum value.
        /// </remarks>
        public string VersionString
        {
            get { return this.versionString; }
            internal set { this.versionString = value; }
        }

        /// <summary>
        /// Override ToString method
        /// </summary>
        /// <returns>Canonical ExchangeService version string</returns>
        public override string ToString()
        {
            return String.Format(
                                 "{0:d}.{1:d2}.{2:d4}.{3:d3}",
                                 this.MajorVersion,
                                 this.MinorVersion,
                                 this.MajorBuildNumber,
                                 this.MinorBuildNumber);
        }
    }
}