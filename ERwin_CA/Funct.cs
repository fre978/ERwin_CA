using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class Funct
    {
        public static bool ParseDataType(string value, string databaseType)
        {
            string[] actualDB = null;
            if (!ConfigFile.DBS.Contains(databaseType))
                return false;
            else
            {
                switch (databaseType)
                {
                    case ConfigFile.DB2_NAME:
                        actualDB = ConfigFile.DATATYPE_DB2;
                        break;
                    case ConfigFile.ORACLE:
                        actualDB = ConfigFile.DATATYPE_ORACLE;
                        break;
                    case ConfigFile.SQLSERVER:
                        break;
                }
            }
            int oUt1;
            int oUt2;
            if (value.Contains(","))
            {
                try
                {
                    string[] a = value.Split('(');
                    string primo = a[0];
                    string[] b = a[1].Split(',');
                    string secondo = b[0];
                    string[] c = (b[1]).Split(')');
                    string terzo = c[0];
                    if (int.TryParse(secondo, out oUt1) && int.TryParse(terzo, out oUt2) && actualDB.Contains(primo.ToLower()))
                        return true;
                    else
                        return false;
                }
                catch(Exception exp)
                {
                    return false;
                }
            }
            if (value.Contains("("))
            {
                try
                {
                    string[] a = value.Split('(');
                    string primo = a[0];
                    string[] b = a[1].Split(')');
                    string secondo = b[0];
                    if (int.TryParse(secondo, out oUt1) && (actualDB.Contains(primo.ToLower()) || actualDB.Contains(primo.ToLower() + "()")))
                        return true;
                    else
                        return false;
                }
                catch(Exception exp)
                {
                    return false;
                }
            }
            else
            {
                if (actualDB.Contains(value.ToLower()))
                    return true;
                else
                    return false;
            }
        }
    }
}
