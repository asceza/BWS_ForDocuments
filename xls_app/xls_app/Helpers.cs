using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls_app
{
    internal class Helpers
    {
        public string FromListToString(List<string> listString)
        {
            string result="";
            foreach (var item in listString)
            {
                result = result + item + "\n";
            }
            return result;
        }
    }
}
