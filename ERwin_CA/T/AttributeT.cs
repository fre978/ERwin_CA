using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA.T
{
    class AttributeT
    {
        public string SSA { get; set; }
        public string Area { get; set; }
        public string NomeTabellaLegacy { get; set; }
        public string NomeCampoLegacy { get; set; }
        public string DefinizioneCampo { get; set; }
        public string TipologiaTabella { get; set; }
        public string DataType { get; set; }
        public int? Lunghezza { get; set; }
        public int? Decimali { get; set; }
        public int? Chiave { get; set; }
        public string Unique { get; set; }
        public string ChiaveLogica { get; set; }
        public int? MandatoryFlag { get; set; }
        public string Dominio { get; set; }
        public string ProvenienzaDominio { get; set; }
        public string Note { get; set; }
        public string Storica { get; set; }
        public string DatoSensibile { get; set; }

        public AttributeT(string nomeTabellaLegacy, string ssa = null, string area = null,
                          string nomeCampoLegacy = null, string definizioneCampo = null, string tipologiaTabella = null,
                          string dataType = null, int? lunghezza = null, int? decimali = null,
                          int? chiave = null, string unique = null, string chiaveLogica= null,
                          int? mandatoryFlag = null, string dominio = null, string provenienzaDominio = null, 
                          string note = null, string storica = null, string datoSensibile = null )
        {
            NomeTabellaLegacy = nomeTabellaLegacy;
            SSA = ssa;
            Area = area;
            NomeCampoLegacy = nomeCampoLegacy;
            DefinizioneCampo = definizioneCampo;
            TipologiaTabella = tipologiaTabella;
            DataType = dataType;
            Lunghezza = lunghezza;
            Decimali = decimali;
            Chiave = chiave;
            Unique = unique;
            ChiaveLogica = chiaveLogica;
            MandatoryFlag = mandatoryFlag;
            Dominio = dominio;
            ProvenienzaDominio = provenienzaDominio;
            Note = note;
            Storica = storica;
            DatoSensibile = datoSensibile;
        }
    }
}
