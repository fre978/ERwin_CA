using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA.T
{
    class RelationT
    {
        public int Row { get; set; }
        public string DB { get; set; }
        public string IdentificativoRelazione { get; set; }
        public string TabellaPadre { get; set; }
        public string TabellaFiglia { get; set; }
        public int? Cardinalita { get; set; }
        public string CampoPadre { get; set; }
        public string CampoFiglio { get; set; }
        public int? Identificativa { get; set; }
        public string Eccezioni { get; set; }
        public bool? TipoRelazione { get; set; }
        public string Note { get; set; }
        public string History { get; set; }
        public int? NullOptionType { get; set; }


public RelationT( int row, string db, 
    string identificativoRelazione = null, string tabellaPadre = null, string tabellaFiglia = null, int? cardinalita = null, 
    string campoPadre = null, string campoFiglio = null, int? identificativa = null, 
    string eccezioni = null, bool? tipoRelazione = null, string note = null, string history = null, int? nullOptionType = null)
        {
            Row = row;
            DB = db;
            IdentificativoRelazione = identificativoRelazione;
            TabellaPadre = tabellaPadre;
            TabellaFiglia = tabellaFiglia;
            Cardinalita = cardinalita;
            CampoPadre = campoPadre;
            CampoFiglio = campoFiglio;
            Identificativa = identificativa;
            Eccezioni = eccezioni;
            TipoRelazione = tipoRelazione;
            Note = note;
            History = history;
            NullOptionType = nullOptionType;





        }
    }
}
