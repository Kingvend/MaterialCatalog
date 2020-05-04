using CollisionFinder.DB;
using FluentNHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.Mappings
{
    public class MaterialGroupMap : ClassMap<DB.MaterialGroup>
    {
        public MaterialGroupMap()
        {
            Id(x => x.ID);
            Map(x => x.Group_name);
            Map(x => x.Group_code);
            Map(x => x.Group_class_name);

            HasMany(x => x.Material)
                .Inverse()
                .Cascade.All();
        }
    }
}
