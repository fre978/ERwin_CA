using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA.T
{
    class EntityT
    {
        public string TableName { get; set; }
        public string SSA { get; set; }
        public string HostName { get; set; }
        public string DatabaseName { get; set; }
        public string Schema { get; set; }
        public string TableDescr { get; set; }
        public string InfoType { get; set; }
        public string TableLimit { get; set; }
        public string TableGranularity { get; set; }
        public string FlagBFD { get; set; }
        // Seconda release
        public string Acronym { get; set; }
        public string Area { get; set; }
        public string TableType { get; set; }
        public string History { get; set; }


        public EntityT(string tName, string ssa = null, string hName = null,
                        string dbName = null, string schema = null, string tableDescr = null,
                        string infoType = null, string tableLimit = null, 
                        string tableGranularity = null, string flagBFD = null, string acronym = null,
                        string area = null, string tableType = null, string history = null)
        {
            TableName = tName;
            SSA = ssa;
            HostName = hName;
            DatabaseName = dbName;
            Schema = schema;
            TableDescr = tableDescr;
            InfoType = infoType;
            TableLimit = tableLimit;
            TableGranularity = tableGranularity;
            FlagBFD = flagBFD;
            //Seconda release
            Acronym = acronym;
            Area = area;
            TableType = tableType;
            History = history;
        }
    }
}
