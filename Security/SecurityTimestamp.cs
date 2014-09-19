// ---------------------------------------------------------------------------
// <copyright file="SecurityTimestamp.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>
// Defines the SecurityTimestamp class.  Note that this file is
// excerpted from WCF sources (ndp\indigo\src\ServiceModel\
// System\ServiceModel\Security\SecurityTimestamp.cs).
// </summary>
//-----------------------------------------------------------------------

// ********************************************************************************
// * NOTE: As noted above, this file is excerpted directly from the WCF source    *
// * code base.  Please do NOT modify this file so as to make keeping it in sync  *
// * with the WCF sources easier.                                                 *
// ********************************************************************************

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Globalization;
    using System.Xml;

    internal sealed class SecurityTimestamp
    {
        //  Pulled from SecurityProtocolFactory
        //
        internal const string DefaultTimestampValidityDurationString = "00:05:00";
        internal static readonly TimeSpan DefaultTimestampValidityDuration = TimeSpan.Parse(DefaultTimestampValidityDurationString);

        internal const string DefaultFormat = "yyyy-MM-ddTHH:mm:ss.fffZ";

        //                            012345678901234567890123
        internal static readonly TimeSpan DefaultTimeToLive = DefaultTimestampValidityDuration;
        private readonly string id;
        private readonly string digestAlgorithm;
        private readonly byte[] digest;
        private char[] computedCreationTimeUtc;
        private char[] computedExpiryTimeUtc;
        private DateTime creationTimeUtc;
        private DateTime expiryTimeUtc;

        public SecurityTimestamp(DateTime creationTimeUtc, DateTime expiryTimeUtc, string id)
            : this(creationTimeUtc, expiryTimeUtc, id, null, null)
        {
        }

        internal SecurityTimestamp(DateTime creationTimeUtc, DateTime expiryTimeUtc, string id, string digestAlgorithm, byte[] digest)
        {
            EwsUtilities.Assert(
                creationTimeUtc.Kind == DateTimeKind.Utc,
                "SecurityTimestamp.ctor",
                "creation time must be in UTC");
            EwsUtilities.Assert(
                expiryTimeUtc.Kind == DateTimeKind.Utc,
                "SecurityTimestamp.ctor",
                "expiry time must be in UTC");

            if (creationTimeUtc > expiryTimeUtc)
            {
                throw new ArgumentOutOfRangeException("recordedExpiryTime");
            }

            this.creationTimeUtc = creationTimeUtc;
            this.expiryTimeUtc = expiryTimeUtc;
            this.id = id;

            this.digestAlgorithm = digestAlgorithm;
            this.digest = digest;
        }

        public DateTime CreationTimeUtc
        {
            get
            {
                return this.creationTimeUtc;
            }
        }

        public DateTime ExpiryTimeUtc
        {
            get
            {
                return this.expiryTimeUtc;
            }
        }

        public string Id
        {
            get
            {
                return this.id;
            }
        }

        public string DigestAlgorithm
        {
            get
            {
                return this.digestAlgorithm;
            }
        }

        internal byte[] GetDigest()
        {
            return this.digest;
        }

        internal char[] GetCreationTimeChars()
        {
            if (this.computedCreationTimeUtc == null)
            {
                this.computedCreationTimeUtc = ToChars(ref this.creationTimeUtc);
            }
            return this.computedCreationTimeUtc;
        }

        internal char[] GetExpiryTimeChars()
        {
            if (this.computedExpiryTimeUtc == null)
            {
                this.computedExpiryTimeUtc = ToChars(ref this.expiryTimeUtc);
            }
            return this.computedExpiryTimeUtc;
        }

        private static char[] ToChars(ref DateTime utcTime)
        {
            char[] buffer = new char[DefaultFormat.Length];
            int offset = 0;

            ToChars(utcTime.Year, buffer, ref offset, 4);
            buffer[offset++] = '-';

            ToChars(utcTime.Month, buffer, ref offset, 2);
            buffer[offset++] = '-';

            ToChars(utcTime.Day, buffer, ref offset, 2);
            buffer[offset++] = 'T';

            ToChars(utcTime.Hour, buffer, ref offset, 2);
            buffer[offset++] = ':';

            ToChars(utcTime.Minute, buffer, ref offset, 2);
            buffer[offset++] = ':';

            ToChars(utcTime.Second, buffer, ref offset, 2);
            buffer[offset++] = '.';

            ToChars(utcTime.Millisecond, buffer, ref offset, 3);
            buffer[offset++] = 'Z';

            return buffer;
        }

        private static void ToChars(int n, char[] buffer, ref int offset, int count)
        {
            for (int i = offset + count - 1; i >= offset; i--)
            {
                buffer[i] = (char) ('0' + (n % 10));
                n /= 10;
            }
            EwsUtilities.Assert(
                n == 0,
                "SecurityTimestamp.ToChars",
                "Overflow in encoding timestamp field");
            offset += count;
        }

        public override string ToString()
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "SecurityTimestamp: Id={0}, CreationTimeUtc={1}, ExpirationTimeUtc={2}",
                this.Id,
                XmlConvert.ToString(this.CreationTimeUtc, XmlDateTimeSerializationMode.RoundtripKind),
                XmlConvert.ToString(this.ExpiryTimeUtc, XmlDateTimeSerializationMode.RoundtripKind));
        }
    }
}
