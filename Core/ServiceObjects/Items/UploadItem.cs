using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    public class UploadItem : ServiceObject
    {
        public UploadItem(ExchangeService service) : base(service) { }

        internal override ServiceObjectSchema GetSchema()
        {
            return UploadSchema.Instance;
        }

        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        internal override void InternalLoad(PropertySet propertySet)
        {
            throw new NotImplementedException();
        }

        internal override void InternalDelete(DeleteMode deleteMode, SendCancellationsMode? sendCancellationsMode, AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            throw new NotImplementedException();
        }

        internal override PropertyDefinition GetIdPropertyDefinition()
        {
            return UploadSchema.Id;
        }

        public ItemId Id
        {
            get { return (ItemId)this.PropertyBag[GetIdPropertyDefinition()]; }
            set { this.PropertyBag[GetIdPropertyDefinition()] = value; }
        }

        public FolderId ParentFolderId
        {
            get { return (FolderId)this.PropertyBag[UploadSchema.ParentFolderId]; }
            set { this.PropertyBag[UploadSchema.ParentFolderId] = value; }
        }

        public byte[] Data
        {
            get { return (byte[])this.PropertyBag[UploadSchema.Data]; }
            set { this.PropertyBag[UploadSchema.Data] = value; }
        }

        public CreateAction CreateAction { get; set; }
    }
}
