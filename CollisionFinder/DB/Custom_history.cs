using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    class Custom_history
    {
        public int ID { get; set; }
        public int Material_ID { get; set; }
        public string Consignee_detail { get; set; }
        public string Shipment_date { get; set; }
        public string Basis_measure_unit { get; set; }
        public double Count_BMU { get; set; }
        public double Shipment_price_BMU { get; set; }
        public string Alt_measure_unit { get; set; }
        public double Count_AMU { get; set; }
        public double Shipment_price_AMU { get; set; }

        public DB.Material Material { get; set; }
    }
}
