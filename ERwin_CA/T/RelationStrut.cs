using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class RelationStrut
    {
        public string ID { get; set;}
        public List<RelationT> Relazioni { get; set; }

        public RelationStrut()
        {
            Relazioni = new List<RelationT>();
        }
    }
}
