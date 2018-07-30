using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response from an Export Items Request
    /// </summary>
    public class ExportItemsResponse : ServiceResponse
    {
        private ItemId itemId = new ItemId();
        private byte[] data;

        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            if(!reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ItemId))
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ItemId);
            }

            itemId.LoadFromXml(reader, XmlNamespace.Messages, XmlAttributeNames.ItemId);

            if(!reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.Data))
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Data);
            }
            data = reader.ReadBase64ElementValue();
        }
        
        /// <summary>
        /// ItemId of the item exported.
        /// </summary>

        public ItemId ItemId
        {
            get
            {
                return itemId;
            }
        }

        /// <summary>
        /// Field containing the item exported.
        /// </summary>
        public byte[] Data
        {
            get
            {
                return data;
            }
        }

    }
}
