using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA.T
{
    class GlobalRelationStrut
    {
        public List<RelationStrut> GlobalRelazioni { get; set; }

        public GlobalRelationStrut()
        {
            GlobalRelazioni = new List<RelationStrut>();
        }
    }
}
