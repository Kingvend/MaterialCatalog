using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    class Material_code
    {
        public int ID { get; set; }
        public int Material_ID { get; set; }
        public string Basic_code { get; set; }
        public string Alternative_code { get; set; }

        public Material Material { get; set; }
    }
}
