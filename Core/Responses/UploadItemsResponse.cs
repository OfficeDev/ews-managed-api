using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    public class UploadItemsResponse : ServiceResponse
    {
        private ItemId itemId = new ItemId();

        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            if(!reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ItemId))
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ItemId);
            }

            itemId.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.ItemId);
        }

        public ItemId Id
        {
            get { return itemId; }
        }
    }
}
