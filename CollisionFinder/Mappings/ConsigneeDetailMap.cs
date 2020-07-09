using FluentNHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.Mappings
{
    public class ConsigneeDetailMap : ClassMap<DB.ConsigneeDetail>
    {

        public ConsigneeDetailMap()
        {
            Id(x => x.ID);
            Map(x => x.Address);
            HasMany(x => x.CustomHistory)
                            .Inverse()
                            .Cascade.All();
        }
    }
}


