using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder.DB
{
    public class MaterialCode
    {
        public virtual int ID { get; protected set; }
        //public virtual int Material_ID { get; set; }
        public virtual string Basic_code { get; set; }
        public virtual string Alternative_code { get; set; }

        public virtual Material Material { get; set; }
    }
}
