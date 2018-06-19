using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NG2LoadExcelSheets.CommonClass
{
    public class DBConnection
    {
        public string ConnectionString()
        {
            return ConfigurationManager.AppSettings["DBConnection"].ToString().Trim();
        }

        public string MDFLocation()
        {
            return ConfigurationManager.AppSettings["MDFPath"].ToString().Trim();
        }

        // Get sheet name
        public string AllTicketsCreated()
        {
            return ConfigurationManager.AppSettings["AllTicketsCreated"].ToString().Trim();
        }
    }
}
