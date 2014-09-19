// ---------------------------------------------------------------------------
// <copyright file="ChangeCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ChangeCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a collection of changes as returned by a synchronization operation.
    /// </summary>
    /// <typeparam name="TChange">Type representing the type of change (e.g. FolderChange or ItemChange)</typeparam>
    public sealed class ChangeCollection<TChange> : IEnumerable<TChange>
        where TChange : Change
    {
        private List<TChange> changes = new List<TChange>();
        private string syncState;
        private bool moreChangesAvailable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChangeCollection&lt;TChange&gt;"/> class.
        /// </summary>
        internal ChangeCollection()
        {
        }

        /// <summary>
        /// Adds the specified change.
        /// </summary>
        /// <param name="change">The change.</param>
        internal void Add(TChange change)
        {
            EwsUtilities.Assert(
                change != null,
                "ChangeList.Add",
                "change is null");

            this.changes.Add(change);
        }

        /// <summary>
        /// Gets the number of changes in the collection.
        /// </summary>
        public int Count
        {
            get { return this.changes.Count; }
        }

        /// <summary>
        /// Gets an individual change from the change collection.
        /// </summary>
        /// <param name="index">Zero-based index.</param>
        /// <returns>An single change.</returns>
        public TChange this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.changes[index];
            }
        }

        /// <summary>
        /// Gets the SyncState blob returned by a synchronization operation.
        /// </summary>
        public string SyncState
        {
            get { return this.syncState; }
            internal set { this.syncState = value; }
        }

        /// <summary>
        /// Gets a value indicating whether the there are more changes to be synchronized from the server.
        /// </summary>
        public bool MoreChangesAvailable
        {
            get { return this.moreChangesAvailable; }
            internal set { this.moreChangesAvailable = value; }
        }

        #region IEnumerable<TChange> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TChange> GetEnumerator()
        {
            return this.changes.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.changes.GetEnumerator();
        }

        #endregion
    }
}
