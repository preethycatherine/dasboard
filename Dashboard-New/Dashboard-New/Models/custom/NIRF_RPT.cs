using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class NIRF_RPT
    {

        /* public string YEAR { get; set; }
          public int sdm { get; set; }
          public string NPRNO { get; set; }
          public string COOR_NAME { get; set; }
          public string AGENCY { get; set; }
          public string TITLE { get; set; }
          public string SANCTNNO { get; set; }
          public Nullable<System.DateTime> SANCTDTE { get; set; }
          public double amount { get; set; }*/

        public string YEAR { get; set; }
        public int MONTH { get; set; }
        public string NPRNO { get; set; }
        public string SANCTNNO { get; set; }
        public string COOR_NAME { get; set; }
        public string COOR_NAME1 { get; set; }

        public string TITLE { get; set; }
        public string AGENCY { get; set; }
        public string CPRNO { get; set; }
        public string CONS_TYPE { get; set; }
        public string DEPT_CODE { get; set; }
        public string AGENCY_CODE { get; set; }
        public Nullable<System.DateTime> SANCTDTE { get; set; }
        public Nullable<System.DateTime> DATE { get; set; }
        public double AMOUNT { get; set; }
        public double rt { get; set; }
    }
}