using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class chqdrawn
    {
        public string from_year { get; set; }
        public string to_year { get; set; }
        public string from_dt { get; set; }
        public string to_dt { get; set; }
        public string PARTY { get; set; }
        public DateTime? CQDATE { get; set; }
        public string CHEQ_NO { get; set; }
        public double? RSAMT { get; set; }
        public string VOUCHNO { get; set; }
        public string BRNO { get; set; }
        public string NPRNO { get; set; }
        public string CHEK { get; set; }

        public static List<chqdrawn> getUsers()
        {
            List<chqdrawn> users = new List<chqdrawn>()
                {
                     new chqdrawn (){ PARTY="Jon", CHEQ_NO="3455" },
                     new chqdrawn (){  PARTY="Alex", CHEQ_NO="345345" }               
                };
            return users;
        }
    }
}