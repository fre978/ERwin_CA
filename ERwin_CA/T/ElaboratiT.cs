using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ERwin_CA.T
{
    class ElaboratiT
    {
        public string FileElaborato { get; set; }
        public List<EntityT> EntityElaborate { get; set; }

        public ElaboratiT(string fileElaborato, List<EntityT> entityElaborate)
        {
            FileElaborato = fileElaborato;
            EntityElaborate = entityElaborate;
        }
    }
}
