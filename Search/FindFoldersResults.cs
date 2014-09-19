// ---------------------------------------------------------------------------
// <copyright file="FindFoldersResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindFoldersResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the results of a folder search operation.
    /// </summary>
    public sealed class FindFoldersResults : IEnumerable<Folder>
    {
        private int totalCount;
        private int? nextPageOffset;
        private bool moreAvailable;
        private Collection<Folder> folders = new Collection<Folder>();

        /// <summary>
        /// Initializes a new instance of the <see cref="FindFoldersResults"/> class.
        /// </summary>
        internal FindFoldersResults()
        {
        }

        /// <summary>
        /// Gets the total number of folders matching the search criteria available in the searched folder.
        /// </summary>
        public int TotalCount
        {
            get { return this.totalCount; }
            internal set { this.totalCount = value; }
        }

        /// <summary>
        /// Gets the offset that should be used with FolderView to retrieve the next page of folders in a FindFolders operation.
        /// </summary>
        public int? NextPageOffset
        {
            get { return this.nextPageOffset; }
            internal set { this.nextPageOffset = value; }
        }

        /// <summary>
        /// Gets a value indicating whether more folders matching the search criteria.
        /// are available in the searched folder. 
        /// </summary>
        public bool MoreAvailable
        {
            get { return this.moreAvailable; }
            internal set { this.moreAvailable = value; }
        }

        /// <summary>
        /// Gets a collection containing the folders that were found by the search operation.
        /// </summary>
        public Collection<Folder> Folders
        {
            get { return this.folders; }
        }

        #region IEnumerable<Folder> Members

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.Collections.Generic.IEnumerator`1"/> that can be used to iterate through the collection.
        /// </returns>
        public IEnumerator<Folder> GetEnumerator()
        {
            return this.folders.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.
        /// </returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.folders.GetEnumerator();
        }

        #endregion
    }
}