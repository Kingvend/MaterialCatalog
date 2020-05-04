using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    public class Material
    {
        public virtual int ID { get; protected set; }
        public virtual string Basic_code { get; set; }
        public virtual string IsHide { get; set; }
        public virtual string Material_name { get; set; }
        public virtual string Material_fullname { get; set; }
        //public virtual int Material_group_ID { get; set; }
        public virtual string Measure_unit { get; set; }

        public virtual IList<CustomHistory> CustomHistory { get; set; }
        public virtual IList<MaterialCode> MaterialCode { get; set; }

        public virtual MaterialGroup MaterialGroup { get; set; }

        public Material()
        {
            CustomHistory = new List<CustomHistory>();
            MaterialCode = new List<MaterialCode>();
        }
    }

    
}
