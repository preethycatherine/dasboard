using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class chqdrawn
    {
        public int PageIndex { get; set; }
        public int PageSize { get; set; }
        public int RecordCount { get; set; }

        public string PARTY { get; set; }
        public DateTime? CQDATE { get; set; }
        public string CHEQ_NO { get; set; }
        public double? RSAMT { get; set; }
        public string VOUCHNO { get; set; }
        public string BRNO { get; set; }
        public string NPRNO { get; set; }
        public string CHEK { get; set; }
        

        public string PARTY_dup { get; set; }
        public DateTime? CQDATE_dup { get; set; }
        public string CHEQ_NO_dup { get; set; }
        public double? RSAMT_dup { get; set; }
        public string VOUCHNO_dup { get; set; }
        public string BRNO_dup { get; set; }
        public string NPRNO_dup { get; set; }
        public string CHEK_dup { get; set; }

        public chqdrawn()
        {
            PARTY_dup = string.Empty;
            CQDATE_dup = DateTime.Now;
        CHEQ_NO_dup = string.Empty;
            RSAMT_dup = 0;
            VOUCHNO_dup = string.Empty;
            BRNO_dup = string.Empty;
            NPRNO_dup = string.Empty;
            CHEK_dup = string.Empty;
        }

    }
}