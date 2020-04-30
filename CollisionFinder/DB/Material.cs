using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    class Material
    {
        public int ID { get; set; }
        public string Basic_code { get; set; }
        public string IsHide { get; set; }
        public string Material_name { get; set; }
        public string Material_fullname { get; set; }
        public int Material_group_ID { get; set; }
        public string Measure_unit { get; set; }

        public List<Custom_history> CustomHistory { get; set; }
        public List<Material_code> MaterialCode { get; set; }

        public MaterialGroup MaterialGroup { get; set; }
    }
}
