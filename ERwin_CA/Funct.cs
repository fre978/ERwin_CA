using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERwin_CA
{
    class Funct
    {
        public static string[] ParseDataType(string value)
        {
            if (value.Contains(","))
            {
                try
                {
                    string[] a = value.Split("(");
                    string primo = a[0];
                    string[] b = a.Split(",");
                    string secondo = b[0];
                    string[] c = (b[1]).Split(")");
                    string terzo = c[0];
                    return new string[] { primo, secondo, terzo };
                }
                catch(Exception exp)
                {
                    return null;
                }
            }
            if (value.Contains("("))
            {
                try
                {
                    string[] a = value.Split("(");
                    string primo = a[0];
                    string[] b = a[1].Split(")");
                    string secondo = b[0];
                    return new string[] { primo, secondo };
                }
                catch(Exception exp)
                {
                    return null;
                }
            }
            else
            {
                string[] a = new string[] { value };
                return a;
            }
        }
    }
}
