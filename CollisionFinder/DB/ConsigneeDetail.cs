using System.Collections.Generic;

namespace CollisionFinder.DB
{
    public class ConsigneeDetail
    {
        public virtual int ID { get; protected set; }
        public virtual string Address { get; set; }

        public virtual IList<DB.CustomHistory> CustomHistory {get; set;}

        public ConsigneeDetail()
        {
            CustomHistory = new List<CustomHistory>();
        }
    }

    
}
