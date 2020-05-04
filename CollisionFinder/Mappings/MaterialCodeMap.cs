using FluentNHibernate.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.Mappings
{
    public class MaterialCodeMap : ClassMap<DB.MaterialCode>
    {
        public MaterialCodeMap()
        {
            Id(x => x.ID);
            //Map(x => x.Material_ID);
            Map(x => x.Basic_code);
            Map(x => x.Alternative_code);

            References(x => x.Material);
        }
    }
}
