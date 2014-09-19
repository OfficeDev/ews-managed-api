// ---------------------------------------------------------------------------
// <copyright file="ExchangeResourceManager.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Exchange Resource Manager.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Reflection;
    using System.Resources;

    /// <summary>
    /// Exchange Resource Manager.
    /// </summary>
    /// <remarks>
    /// The Exchange Resource Manager gives us access to the assembly name.
    /// This allows the LocalizedString to try to reconstruct a "serialized"
    /// resource manager in the client side. If the client does not have
    /// the corresponding assembly, the resource manager will not be constructed,
    /// of course. See the description in LocalizedString for more details.
    /// </remarks>
    internal sealed class ExchangeResourceManager : ResourceManager
    {
        // Dictionary of resource managers. Initialized only if someone uses resources in the process.
        private static System.Collections.Specialized.HybridDictionary resourceManagers = new System.Collections.Specialized.HybridDictionary();

        /// <summary>
        /// lock object used when accessing ResourceManager
        /// </summary>
        private static object lockObject = new object();

        /// <summary>
        /// Returns the instance of the ExchangeResourceManager class that looks up 
        /// resources contained in files derived from the specified root name using the given Assembly.
        /// <see cref="System.Resources.ResourceManager"/>
        /// </summary>
        /// <param name="baseName">The root name of the resources.</param>
        /// <param name="assembly">The main Assembly for the resources.</param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="assembly"/> is null.
        /// </exception>
        /// <returns>ExchangeResourceManager</returns>
        public static ExchangeResourceManager GetResourceManager(string baseName, Assembly assembly)
        {
            if (null == assembly)
            {
                throw new ArgumentNullException("assembly");
            }

            string key = baseName + assembly.GetName().Name;

            lock (lockObject)
            {
                ExchangeResourceManager resourceManager = (ExchangeResourceManager)resourceManagers[key];
                if (null == resourceManager)
                {
                    resourceManager = new ExchangeResourceManager(baseName, assembly);
                    resourceManagers[key] = resourceManager;
                }
                return resourceManager;
            }
        }

        /// <summary>
        /// Creates a new instance of this class.
        /// </summary>
        /// <param name="baseName">The root name of the resources.</param>
        /// <param name="assembly">The main Assembly for the resources.</param>
        private ExchangeResourceManager(string baseName, Assembly assembly)
            : base(baseName, assembly)
        {
        }

        /// <summary>
        /// Base Name for the resources
        /// </summary>
        /// <remarks>
        /// Used by LocalizedString to serialize localized strings.
        /// </remarks>
        public override string BaseName
        {
            get { return base.BaseName; }
        }

        /// <summary>
        /// Assembly containing the resources
        /// </summary>
        /// <remarks>
        /// Used by LocalizedString to serialize localized strings.
        /// </remarks>
        public string AssemblyName
        {
            // NOTE: do we want to use the full name? What if the client is a Service Pack off?
            get { return MainAssembly.GetName().Name; }
        }

        /// <summary>
        /// Retrieves a string from the resource table based on a string id.
        /// Asserts if the string cannot be found.
        /// </summary>
        /// <param name="name">Id of the string to retrieve.</param>
        /// <returns>The corresponding string if the id was located in the table, null otherwise.</returns>
        public override string GetString(string name)
        {
            return this.GetString(name, System.Globalization.CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Retrieves a string from the resource table based on a string id.
        /// Asserts if the string cannot be found.
        /// </summary>
        /// <param name="name">Id of the string to retrieve.</param>
        /// <param name="culture">The culture to use.</param>
        /// <returns>The corresponding string if the id was located in the table, null otherwise.</returns>
        public override string GetString(string name, System.Globalization.CultureInfo culture)
        {
            return base.GetString(name, culture);
        }
    }
}
