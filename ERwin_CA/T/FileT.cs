using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class FileT
    {
        public string SSA { get; set; }
        public string Acronimo { get; set; }
        public string NomeModello { get; set; }
        public string TipoDBMS { get; set; }
        public string Estensione { get; set; }

        public FileT(string ssa = null, string acronimo = null, string nomemodello = null, string tipodbms = null, string estensione = null)
        {
            SSA = ssa;
            Acronimo = acronimo;
            NomeModello = nomemodello;
            TipoDBMS = tipodbms;
            Estensione = estensione;
        }
    }
}
