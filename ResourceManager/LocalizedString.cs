// ---------------------------------------------------------------------------
// <copyright file="LocalizedString.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Struct that defines a localized string.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Globalization;
    using System.Reflection;
    using System.Runtime.Serialization;

    /// <summary>
    /// Interface implemented by objects that provide a LocalizedString.
    /// </summary>
    internal interface ILocalizedString
    {
        /// <summary>
        /// LocalizedString held by this object.
        /// </summary>
        LocalizedString LocalizedString { get; }
    }

    /// <summary>
    /// Struct that defines a localized string.
    /// </summary>
    [Serializable]
    internal struct LocalizedString : ISerializable, ILocalizedString, IFormattable, IEquatable<LocalizedString>
    {
        /*
            This structure binds together the parameters necessary to move
            a localized string around the system. The resource ID, the
            resource manager and insert objects must stay together so 
            that we'll have all the information necessary to localize 
            the string when a client requests it in any given culture.
            While the data stays in the same remoting boundary everything
            is fine. A problem arises when we need to cross a remoting
            bounduary - the .net ResourceManager is not remotable due
            to the RuntimeResourceSet object. So, the client would not
            get the resource manager at the other end or, if anything, it
            would need a reference back to the server which is bad, since
            it may take forever before the client actually localizes the
            string. To minimize the problem, here's our current solution:
            When we serialize the localized string we also serialize the
            BaseName and AssemblyName of the resource manager. This is
            how ExchangeResourceManagers are constructed so if we can load the 
            assembly on the client we should be able to reconstruct the
            ExchangeResourceManager, without requiring a server object. Note
            that this requires the resource DLL to be on the client - 
            not a hard requirement since you are explicitly sending data
            that should be localized on the client, so that's fine.
            If we get the resource manager on the client, then we are
            happy. If for any reason we cannot recreate the resource
            manager, we'll use a fallback format string - localized with 
            the culture from the server - hey, it's better than showing a
            resource ID. But then again, that will only happen
            if you are not dropping the resource DLL in the client.
        */

        /// <summary>
        /// The id of the localized string.
        /// </summary>
        /// <remarks>
        /// If we don't have a ResourceManager, this is
        /// the formating string we'll use in ToString().
        /// This can happen if we serialize the object and
        /// we are unable to reload the resource manager
        /// when deserializing.
        /// </remarks>
        private readonly string Id;

        /// <summary>
        /// Strings to be inserted in the message identified by Id.
        /// </summary>
        private readonly object[] Inserts;

        /// <summary>
        /// Resource Manager capable of loading the string.
        /// </summary>
        private readonly ExchangeResourceManager ResourceManager;

        /// <summary>
        /// The one and only LocalizedString.Empty.
        /// </summary>
        public static readonly LocalizedString Empty = new LocalizedString();

        /// <summary>
        /// Compares both strings.
        /// </summary>
        /// <param name="s1">First string.</param>
        /// <param name="s2">Second string.</param>
        /// <returns>True if objects are equal.</returns>
        public static bool operator ==(LocalizedString s1, LocalizedString s2)
        {
            return s1.Equals(s2);
        }

        /// <summary>
        /// Compares both strings.
        /// </summary>
        /// <param name="s1">First string.</param>
        /// <param name="s2">Second string.</param>
        /// <returns>True if objects are not equal.</returns>
        public static bool operator !=(LocalizedString s1, LocalizedString s2)
        {
            return !s1.Equals(s2);
        }

        /// <summary>
        /// Implicit conversion from a LocalizedString to a string.
        /// </summary>
        /// <param name="value">LocalizedString value to convert to a string.</param>
        /// <returns>The string localized in the CurrentCulture.</returns>
        /// <remarks>
        /// While the rule of thumb says that an implicit conversion
        /// should not loose data, this operator is an exception.
        /// The moment a LocalizedString becomes a string, we lose
        /// the localization information and we end up with the
        /// localized string in the current culture - from there
        /// we cannot go back to a fully localizable string.
        /// We allow that because the usage pattern of LocalizedString
        /// is so that by the time we convert a LocalizedString to
        /// a string we are about to show the string to the client.
        /// Most certainly we'll never import that string back
        /// into a LocalizedString again, so it really does not matter
        /// that we're loosing the information.
        /// </remarks>
        public static implicit operator string(LocalizedString value)
        {
            return value.ToString();
        }

        /// <summary>
        /// Joins objects in a localized string.
        /// </summary>
        /// <param name="separator">Separator between strings.</param>
        /// <param name="value">Array of objects to join as strings.</param>
        /// <returns>
        /// A LocalizedString that concatenates the given objects.
        /// </returns>
        public static LocalizedString Join(object separator, object[] value)
        {
            if (null == value || 0 == value.Length)
            {
                return LocalizedString.Empty;
            }

            if (null == separator)
            {
                separator = string.Empty;
            }

            // Create the insert array, containing the separator (which can also be localized)
            // and the original parameters.
            object[] insert = new object[value.Length + 1];
            insert[0] = separator;
            Array.Copy(value, 0, insert, 1, value.Length);

            // Build a format string that will concatenate all objects
            System.Text.StringBuilder sb = new System.Text.StringBuilder(6 * value.Length);
            sb.Append("{");
            for (int n = 1; n < value.Length; n++)
            {
                sb.Append(n);
                sb.Append("}{0}{");
            }
            sb.Append(value.Length + "}");

            return new LocalizedString(sb.ToString(), insert);
        }

        /// <summary>
        /// Creates a new instance of the structure.
        /// </summary>
        /// <param name="id">The id of the localized string.</param>
        /// <param name="resourceManager">Resource Manager capable of loading the string.</param>
        /// <param name="inserts">Strings to be inserted in the message identified by Id.</param>
        public LocalizedString(string id, ExchangeResourceManager resourceManager, params object[] inserts)
        {
            if (null == id)
            {
                throw new ArgumentNullException("id");
            }

            if (null == resourceManager)
            {
                throw new ArgumentNullException("resourceManager");
            }

            this.Id = id;
            this.ResourceManager = resourceManager;

            // If no inserts are passed, inserts is object[0] rather than null
            this.Inserts = ((inserts != null) && (inserts.Length > 0)) ? inserts : null;
        }

        /// <summary>
        /// Encapsulates a string in a LocalizedString.
        /// </summary>
        /// <remarks>
        /// While the rule of thumb says that an implicit conversion
        /// can be used when there's no loss of data, this is not the case 
        /// with this constructor. When going from string to LocalizedString
        /// we don't lose information but we don't gain information
        /// either. The usage pattern of LocalizedString asks that
        /// if a string is to be localizable it should always be
        /// transported around in a LocalizedString. If you are setting
        /// a LocalizedString from a string it is most likely
        /// that you lost data already, somewhere else. To flag this
        /// problem, instead of an implicit conversion we have a constructor
        /// to remind people that this is not your ideal situation. This way 
        /// we can also search for "new LocalizedString" in the code and 
        /// see where we're doing this and come up with a design where
        /// we will not lose the localization information until it's
        /// time to show the string to the user.
        /// Ideally, we would be able to remove all instances where we
        /// need this constructor, but then people would just create a 
        /// localized string "{0}", which would give us just about the 
        /// same thing with less perf.
        /// </remarks>
        /// <param name="value">
        /// String to encapsulate.
        /// Note that if value is null this creates a copy of 
        /// LocalizedString.Empty and ToString will return "", not null.
        /// This is intentional to avoid returning null from ToString().
        /// </param>
        public LocalizedString(string value)
        {
            this.Id = value;
            this.Inserts = null;
            this.ResourceManager = null;
        }

        /// <summary>
        /// Encapsulates a hardcoded formatting string and 
        /// its parameters as a LocalizedString.
        /// </summary>
        /// <param name="format">Formatting string.</param>
        /// <param name="inserts">Insert parameters.</param>
        /// <remarks>
        /// The formatting string is localized "as-is".
        /// This is used to append strings and other things like that.
        /// </remarks>
        private LocalizedString(string format, object[] inserts)
        {
            this.Id = format;
            this.Inserts = inserts;
            this.ResourceManager = null;
        }

        /// <summary>
        /// Serialization-required constructor 
        /// </summary>
        /// <param name="info">Holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">Contains contextual information about the source or destination.</param>
        private LocalizedString(SerializationInfo info, StreamingContext context)
        {
            this.Inserts = (object[])info.GetValue("inserts", typeof(object[]));

            // The original assembly where the string came from may not exist on the client.
            // In such case, deserializing the entire ExchangeResourceManager would fail since we
            // would not be able to load the Assembly. So we'll try that now and if it fails,
            // too bad...we'll just use the fallback string when formatting the string.
            this.ResourceManager = null;
            this.Id = null;
            try
            {
                string baseName = info.GetString("baseName");
                string assemblyName = info.GetString("assemblyName");
                Assembly assembly = Assembly.Load(assemblyName);
                this.ResourceManager = ExchangeResourceManager.GetResourceManager(baseName, assembly);
                this.Id = info.GetString("id");
                if (null == this.ResourceManager.GetString(this.Id))    // to check the resource actually exists
                {
                    this.ResourceManager = null;                        // The resource manager we got does not contain the string we're looking for.
                }
            }
            catch (System.Runtime.Serialization.SerializationException) // no assembly data
            {
                // make presharp happy (does not like empty catch blocks)
                this.ResourceManager = null;
            }
            catch (System.IO.FileNotFoundException)                     // assembly not found
            {
                // make presharp happy (does not like empty catch blocks)
                this.ResourceManager = null;
            }
            catch (System.Resources.MissingManifestResourceException)   // resource file not found
            {
                this.ResourceManager = null;
            }

            if (null == this.ResourceManager)
            {
                // Well, we don't have a resource manager so there's no point
                // keeping an ID around. Let's load the formatting string in the ID
                // and use that to format the insert parameters.
                this.Id = info.GetString("fallback");
            }
        }

        /// <summary>
        /// Called when the object is serialized. 
        /// </summary>
        /// <remarks>
        /// When serializing the insert parameters we will replace any non-serializable object
        /// with its ToString() version or its ILocalizedString.LocalizedString property.
        /// </remarks>
        /// <param name="info">Holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">Contains contextual information about the source or destination.</param>
        [System.Security.Permissions.SecurityPermissionAttribute(System.Security.Permissions.SecurityAction.LinkDemand, Flags = System.Security.Permissions.SecurityPermissionFlag.SerializationFormatter)]
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            // Look for non-serializable types in the array. If we find one, serialize its string.
            object[] serializableInserts = null;
            if ((null != this.Inserts) && (this.Inserts.Length > 0))
            {
                serializableInserts = new object[this.Inserts.Length];
                for (int i = 0; i < this.Inserts.Length; i++)
                {
                    object serializableInsert = this.Inserts[i];

                    if (null != serializableInsert)
                    {
                        // If the serializable object is ILocalizedString, serialize the property.
                        // Otherwise, serialize just the string representation of the object.
                        if (serializableInsert is ILocalizedString)
                        {
                            serializableInsert = ((ILocalizedString)serializableInsert).LocalizedString;
                        }
                        else if (! serializableInsert.GetType().IsSerializable && !(serializableInsert is ISerializable))
                        {
                            // Since the object is not good for serialization, let's translate it now.
                            // If the translation has nothing better to offer, serialize the object's ToString()
                            // otherwise we'll serialize the translation.
                            object translation = TranslateObject(serializableInsert, CultureInfo.InvariantCulture);
                            if (translation == serializableInsert)
                            {
                                serializableInsert = serializableInsert.ToString();
                            }
                            else
                            {
                                serializableInsert = translation;
                            }
                        }
                    }

                    serializableInserts[i] = serializableInsert;
                }
            }
            info.AddValue("inserts", serializableInserts);

            // This may be null if we were deserialized and unable to get the resource manager back.
            // If that happens once, we'll never try to load the resource manager again.
            // The Id of the string will be the formatting string when the resource manager is null.
            // While we have the resource manager, save everything we need to try to reload it.
            if (null != this.ResourceManager)
            {
                info.AddValue("baseName", this.ResourceManager.BaseName);
                info.AddValue("assemblyName", this.ResourceManager.AssemblyName);
                info.AddValue("id", this.Id);
                info.AddValue("fallback", this.ResourceManager.GetString(this.Id, System.Globalization.CultureInfo.InvariantCulture));
            }
            else
            {
                // Tough luck - we don't know if this is a good language for the client, but what can we do?!?
                info.AddValue("fallback", this.Id);
            }
        }

        /// <summary>
        /// Returns the string localized in the current UI culture.
        /// </summary>
        /// <returns>The localized string.</returns>
        public override string ToString()
        {
            return ((IFormattable)this).ToString(null, null);
        }

        /// <summary>
        /// Returns the string localized in the given culture.
        /// </summary>
        /// <param name="formatProvider">
        /// The <see cref="IFormatProvider"/> to use to format the value or
        /// a <see langword="null"/> reference to obtain the format information 
        /// from the current UI culture. This parameter is usually a 
        /// <see cref="CultureInfo"/> object.
        /// </param>
        /// <returns>The localized string.</returns>
        /// <remarks>
        /// Note that neutral cultures are unable to format
        /// strings that contain numeric or date/time insertion parameters.
        /// </remarks>
        public string ToString(IFormatProvider formatProvider)
        {
            return ((IFormattable)this).ToString(null, formatProvider);
        }

        /// <summary>
        /// Returns the string localized in the given culture.
        /// </summary>
        /// <param name="format">
        /// The <see cref="string"/> specifying the format to use or 
        /// a <see langword="null"/> reference to use the default format 
        /// defined for the type of the <see cref="IFormattable"/> implementation. 
        /// This parameter is currently ignored.
        /// </param>
        /// <param name="formatProvider">
        /// The <see cref="IFormatProvider"/> to use to format the value or
        /// a <see langword="null"/> reference to obtain the format information 
        /// from the current UI culture. 
        /// If this parameter is a <see cref="CultureInfo"/> the resulting
        /// string will be localized in the given culture otherwise the
        /// current UI culture will be used to load the string from the
        /// resource file.
        /// </param>
        /// <returns>The string localized in the given culture.</returns>
        string IFormattable.ToString(string format, IFormatProvider formatProvider)
        {
            if (this.IsEmpty)
            {
                return string.Empty;
            }

            // If the resource manager is set, then ID is the string ID of the formatting string, so let's
            // get the one that is good for the language we're localizing to. Otherwise, the ID is the last
            // formatting string we were able to grab. A side effect of this is that the message may end
            // up partially localized - is that better than something that the user can not understand a word
            // or not?
            format = (null != this.ResourceManager) ?
                this.ResourceManager.GetString(this.Id, formatProvider as CultureInfo) :
                this.Id;

            // Look for any "bad objects" in the insert parameters and
            // pick a better one. Otherwise, things like LocalizedException
            // will return the ToString() instead of the intended string.
            // We will also special case some classes that we know don't do 
            // a good job in their ToString() and will replace nulls to
            // prevent formating exceptions. The most special case is when
            // we have another LocalizedString as an insertion parameter.
            // We need to get its localized string in the correct culture!
            if (null != this.Inserts && this.Inserts.Length > 0)  // if Inserts.Length == 0, string.Format throws a FormatException
            {
                object[] insertParams = new object[this.Inserts.Length];
                for (int i = 0; i < this.Inserts.Length; i++)
                {
                    object insertParam = this.Inserts[i];

                    if (insertParam is ILocalizedString)
                    {
                        insertParam = ((ILocalizedString)insertParam).LocalizedString;
                    }
                    else
                    {
                        insertParam = TranslateObject(insertParam, formatProvider); // Don't ToString since the object may be IFormattable.
                    }

                    // Note: If param is null value, just pass it unaltered down to string.Format.
                    // This keeps format string compatible between string.Format and LocalizedString.ToString().
                    insertParams[i] = insertParam;
                }

                return string.Format(formatProvider, format, insertParams);
            }
            else
            {
                return format;
            }
        }

        /// <summary>
        /// Returns the object itself.
        /// </summary>
        LocalizedString ILocalizedString.LocalizedString
        {
            get { return this; }
        }

        /// <summary>
        /// True if the string is empty.
        /// </summary>
        /// <remarks>
        /// This is slighly faster than comparing the string against LocalizedString.Empty.
        /// </remarks>
        public bool IsEmpty
        {
            get { return null == this.Id; }
        }

        /// <summary>
        /// Returns a hash code based on the hash of the resource manager and the hash of the ID.
        /// </summary>
        /// <returns>Hash code of object.</returns>
        public override int GetHashCode()
        {
            int idHash = (null != this.Id) ? this.Id.GetHashCode() : 0;
            int rmHash = (null != this.ResourceManager) ? this.ResourceManager.GetHashCode() : 0;
            int hash = idHash ^ rmHash;
            if (null != this.Inserts)
            {
                for (int i = 0; i < this.Inserts.Length; i++)
                {
                    hash = (hash ^ i) ^ ((null != this.Inserts[i]) ? this.Inserts[i].GetHashCode() : 0);
                }
            }
            return hash;
        }

        /// <summary>
        /// Compares this string with another.
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>Returns true if objects are equal.</returns>
        public override bool Equals(object obj)
        {
            // If not a locstring, not equal.
            if (obj is LocalizedString)
            {
                return this.Equals((LocalizedString)obj);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Compares this string with another.
        /// </summary>
        /// <param name="obj">Object to compare.</param>
        /// <returns>True if LocalizedString objects are equal.</returns>
        public bool Equals(LocalizedString obj)
        {
            // Their IDs must match, their ExchangeResourceManager must match and all the insert parameters must match.
            if (!string.Equals(this.Id, obj.Id, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (null != this.ResourceManager ^ null != obj.ResourceManager)
            {
                return false;
            }

            if (null != this.ResourceManager && !this.ResourceManager.Equals(obj.ResourceManager))
            {
                return false;
            }

            if (null != this.Inserts ^ null != obj.Inserts)
            {
                return false;
            }

            if (null != this.Inserts && null != obj.Inserts)
            {
                if (this.Inserts.Length != obj.Inserts.Length)
                {
                    return false;
                }

                for (int i = 0; i < this.Inserts.Length; i++)
                {
                    if (null != this.Inserts[i] ^ null != obj.Inserts[i])
                    {
                        return false;
                    }

                    if (null != this.Inserts[i] && null != obj.Inserts[i])
                    {
                        if (!this.Inserts[i].Equals(obj.Inserts[i]))
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Given an object that is not good for serialization or that
        /// returns an awful string in its ToString() this function
        /// will return one that we consider better for the user.
        /// </summary>
        /// <param name="badObject">Object that cannot be serialized</param>
        /// <param name="formatProvider">FormatProvider</param>
        /// <returns>A string or LocalizedString to represent the object.</returns>
        private static object TranslateObject(object badObject, IFormatProvider formatProvider)
        {
            Exception badExObject = badObject as Exception;
            if (badExObject != null)
            {
                // TODO: correctly localize framework exceptions that
                // we know will look wrong if we cross culture boundaries.
                // We'll return those as LocalizedString objects.
                return badExObject.Message;
            }

            return badObject;
        }

        /// <summary>
        /// Returns a numeric Id identifying the localized string template without taking the inserts into consideration.
        /// </summary>
        public int BaseId
        {
            get
            {
                string fullId = ((null != this.ResourceManager) ? this.ResourceManager.BaseName : string.Empty) + this.Id;
                return fullId.GetHashCode();
            }
        }
    }
}
