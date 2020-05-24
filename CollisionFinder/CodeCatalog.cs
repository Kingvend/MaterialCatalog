using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder
{
    class CodeCatalog
    {
        public string Name { get; set; }
        public string BaseCode { get; set; }
        public List<string> AltCode { get; set; }
        public string BaseMU { get; set; }
        public string BaseBrutto { get; set; }

        static public string FindBaseCode(IGrouping<string, MTR_Catalog> MC)
        {
            string baseCode = "" ;
           
            return baseCode;

        }
    }
}
