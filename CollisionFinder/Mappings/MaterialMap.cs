using FluentNHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.Mappings
{
    public class MaterialMap : ClassMap<DB.Material>
    {
        public MaterialMap()
        {
            Id(x => x.ID);
            Map(x => x.Basic_code);
            Map(x => x.IsHide);
            Map(x => x.Material_name);
            Map(x => x.Material_fullname);
            Map(x => x.Measure_unit);
            References(x => x.MaterialGroup);
            HasMany(x => x.CustomHistory)
                .Inverse()
                .Cascade.All();
            HasMany(x => x.MaterialCode)
                .Inverse()
                .Cascade.All();
        }
    }
}
