﻿using FluentNHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.Mappings
{
    public class CustomHistoryMap : ClassMap<DB.CustomHistory>
    { 
        public CustomHistoryMap()
        {
            Id(x => x.ID);
            //Map(x => x.Material_ID);
            Map(x => x.Consignee_detail)
                .Length(1000);
            Map(x => x.Shipment_date);
            Map(x => x.Basis_measure_unit);
            //Map(x => x.Count_BMU);
            //Map(x => x.Shipment_price_BMU);
            //Map(x => x.Alt_measure_unit);
            //Map(x => x.Count_AMU);
            //Map(x => x.Shipment_price_AMU);
            Map(x => x.DATE_SCHF);
            Map(x => x.KOL_SCHF);
            Map(x => x.SUM_SCHF);

            References(x => x.Material);
            References(x => x.ConsigneeDetail);
        }
    }
}
