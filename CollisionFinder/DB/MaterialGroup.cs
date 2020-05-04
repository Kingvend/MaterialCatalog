using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    public class MaterialGroup
    {
        public virtual int ID { get; protected set; }
        public virtual string Group_name { get; set; }
        public virtual string Group_code { get; set; }
        public virtual string Group_class_name { get; set; }

        public virtual IList<Material> Material { get; set; }

        public MaterialGroup()
        {
            Material = new List<Material>();
        }


    }
}
