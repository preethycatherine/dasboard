using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class PBVoucher_pfms
    {
        public string from_year { get; set; }
        public string to_year { get; set; }
        public string from_dt { get; set; }
        public string to_dt { get; set; }
        public string BANK_ACCOUNT_NO { get; set; }
        public string IFSC_CODE { get; set; }
        public string BRANCH_NAME { get; set; }
        public string BANK_NAME { get;  set; }
        public string MONTH { get; set; }
        public string VRNO { get; set; }
        public DateTime DATE { get; set; }
        //public DateTime DINP { get; set; }
        public double AMOUNT { get; set; }
        public string NPRNO { get; set; }
        public string PART { get; set; }
       public string HEAD { get; set; }
        public string DISC { get; set; }
        public string DIS { get; set; }
        //public string ICCNO { get; set; }
        public string PONO { get; set; }
        public string COMNO { get; set; }
        public string CQNO { get; set; }
        //public string OPTION { get; set; }
        public string BRNO { get; set; }
        public string NATURE { get; set; }
        //public string CHECK { get; set; }
        public string REGNO { get; set; }
        public string LEDDIS { get; set; }
        public string ECODE { get; set; }

    }
}