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
	using System.Runtime.Serialization;

	/// <summary>
	/// Represents an error that occurs when the account that is being accessed is locked and requires user interaction to be unlocked.
	/// </summary>
	[Serializable]
	public class AccountIsLockedException : ServiceRemoteException
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="AccountIsLockedException"/> class.
		/// </summary>
		/// <param name="message">Error message text.</param>
		/// <param name="accountUnlockUrl">URL for client to visit to unlock account.</param>
		/// <param name="innerException">Inner exception.</param>
		public AccountIsLockedException(string message, Uri accountUnlockUrl, Exception innerException)
			: base(message, innerException)
		{
			this.AccountUnlockUrl = accountUnlockUrl;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="T:Microsoft.Exchange.WebServices.Data.AccountIsLockedException"/> class with serialized data.
		/// </summary>
		/// <param name="info">The object that holds the serialized object data.</param>
		/// <param name="context">The contextual information about the source or destination.</param>
		protected AccountIsLockedException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
			this.AccountUnlockUrl = (Uri)info.GetValue("AccountUnlockUrl", typeof(Uri));
		}

		/// <summary>Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and additional exception information.</summary>
		/// <param name="info">The object that holds the serialized object data. </param>
		/// <param name="context">The contextual information about the source or destination. </param>
		/// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> object is a null reference (Nothing in Visual Basic). </exception>
		public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			EwsUtilities.Assert(info != null, "AccountIsLockedException.GetObjectData", "info is null");

			base.GetObjectData(info, context);

			info.AddValue("AccountUnlockUrl", this.AccountUnlockUrl, typeof(Uri));
		}

		/// <summary>
		/// Gets the URL of a web page where the user can navigate to unlock his or her account.
		/// </summary>
		public Uri AccountUnlockUrl
		{
			get;
			private set;
		}
	}
}