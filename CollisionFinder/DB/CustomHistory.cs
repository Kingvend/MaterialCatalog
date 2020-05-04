using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    public class CustomHistory
    {
        public virtual int ID { get; protected set; }
        //public virtual int Material_ID { get; set; }
        public virtual string Consignee_detail { get; set; }
        public virtual string Shipment_date { get; set; }
        public virtual string Basis_measure_unit { get; set; }
        public virtual double Count_BMU { get; set; }
        public virtual double Shipment_price_BMU { get; set; }
        public virtual string Alt_measure_unit { get; set; }
        public virtual double Count_AMU { get; set; }
        public virtual double Shipment_price_AMU { get; set; }

        public virtual DB.Material Material { get; set; }
    }
}
