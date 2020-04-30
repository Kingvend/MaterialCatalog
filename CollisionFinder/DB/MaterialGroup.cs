using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    class MaterialGroup
    {
        public int ID { get; set; }
        public string Group_name { get; set; }
        public string Group_code { get; set; }
        public string Group_class_name { get; set; }

        public List<Material> Material { get; set; }
    }
}
