using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{    public class vcvendor
    {
        //internal object Database;
        public string from_dt { get; set; }
        //[Required]
        public string to_dt { get; set; }
        //[Required]
        public string from_dt1 { get; set; }
        //[Required]
        public string to_dt1 { get; set; }
        public System.DateTime DINP { get; set; }
        public string VCTRNO { get; set; }
        public string AMOUNT { get; set; }
        public string NPRNO { get; set; }
        public string ICCNO { get; set; }
        public string COMNO { get; set; }
        public string VRNO { get; set; }
        public string BRNO { get; set; }
        public string VPartyCode { get; set; }
        public string ASSTCK { get; set; }
        public string ACCTCK { get; set; }
        public string ACCT1CK { get; set; }
        public string SOCK { get; set; }
        public string DRCK { get; set; }
        public Nullable<System.DateTime> CRDATE { get; set; }
        public string VCTRBNO { get; set; }
        public string LUSER { get; set; }
        public string vname { get; set; }
        //public List<Dashboard_New.Models.custom.vcvendor> vm { get; set; }
        //public List<Dashboard_New.Models.custom.dcvendor> dm { get; set; }
        //public vcvendor()
        //{

        //    vm = new List<Dashboard_New.Models.custom.vcvendor>();
        //    dm = new List<Dashboard_New.Models.custom.dcvendor>();
        //}
    }
}