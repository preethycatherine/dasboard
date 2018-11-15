using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class billcons
    {
        public string id { get; set; }
        public string from_year { get; set; }
        public string to_year { get; set; }
        public string from_dt { get; set; }
        public string to_dt { get; set; }
        public string BRNO { get; set; }
        public DateTime? BRDATE { get; set; }
        public DateTime? CASHDATE { get; set; }
        public string BILLNO_DT { get; set; }
        public string NPRNO { get; set; }
        public string HEAD { get; set; }
        public string COMNO { get; set; }
        public string PONO { get; set; }
        public string VOCHNO { get; set; }
        public double VAMOUNT { get; set; }
        public string ITEM { get; set; }
        public string PARTY { get; set; }
        public string ADD1 { get; set; }
        public string ADD2 { get; set; }
        public string ADD3 { get; set; }
        public string CITY { get; set; }
        public string NARATION1 { get; set; }
        public string NARATION2 { get; set; }
        public string SPNSTOCSH { get; set; }
        public string SPNSTOCSH1 { get; set; }
        public string SPNSTOCSH2 { get; set; }
        public string CQCK { get; set; }
        public string LEDGER1 { get; set; }
        public string LEDGER { get; set; }
        public string LEDGER2 { get; set; }
        public string SCHOLAR { get; set; }
        public string TIME { get; set; }
        public string ADJ { get; set; }
        public string ECODE { get; set; }
        public double? IT { get; set; }
        public double? PT { get; set; }
        public string PARTYCODE { get; set; }
        public string PANNUMBER { get; set; }
        public string TPNO { get; set; }
        public DateTime? TPLDATE { get; set; }
        public string ACCOUNTTYPE { get; set; }
        public string TDSPERSENT { get; set; }
        public string TDSSECTION { get; set; }
        public decimal? TDSAMOUNT { get; set; }
        public decimal? PARTYAMOUNT { get; set; }
        public decimal? TAXABLEAMOUNT { get; set; }
        public decimal? INVOICEAMOUNT { get; set; }
        public string VOUCHERNUMBER { get; set; }
        public DateTime? VOUCHERDATE { get; set; }
        public decimal? TDSGSTVALUE { get; set; }
        public decimal? TDSGSTAMOUNT { get; set; }
        public Int32? SLNO { get; set; }
        public string TPLNO { get; set; }
    }
}