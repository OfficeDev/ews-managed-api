using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Reflection;

namespace EwsManagedApiTest
{
    /// <summary>
    /// A class for extension methods to assist.
    /// </summary>
    static class ExtensionMethods
    {
        /// <summary>
        /// Returns the collection of TimeZoneTransitions that are contained within the TimeZoneDefinition object.
        /// Uses reflection to access the private field member.
        /// </summary>
        /// <param name="timeZoneDefinition">The TimeZoneDefinition object whose transitions are to be returned.</param>
        /// <returns>The collection of TimeZoneTransition objects.</returns>
        public static List<TimeZoneTransition> GetTransitions(this TimeZoneDefinition timeZoneDefinition)
        {
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            FieldInfo field = typeof(TimeZoneDefinition).GetField("transitions", bindFlags);
            return field.GetValue(timeZoneDefinition) as List<TimeZoneTransition>;
        }
    }
}
