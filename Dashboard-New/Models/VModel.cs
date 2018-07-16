using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models
{
    public class VModel
    {
        //[Required(ErrorMessage = "Date is required")]
        public string from_dt { get; set; }
        //[Required]
        public string to_dt { get; set; }
        //[Required]
        public string from_dt1 { get; set; }
        //[Required]
        public string to_dt1 { get; set; }

        public string from_year { get; set; }
        public string to_year { get; set; }
        
        public List<VCVOUCH> vclist { get; set; }
        public List<DVOUCH> dclist { get; set; }

        public List<MSTLST> mstlist { get; set; }
        public List<REC1718> reclist { get; set; }
        public List<Dashboard_New.Models.custom.vcvendor> vm { get; set; }
        public List<Dashboard_New.Models.custom.dcvendor> dm { get; set; }
        public List<Dashboard_New.Models.custom.NIRF_RPT> nirf { get; set; }
        public List<Dashboard_New.Models.custom.ppo> ppoo { get; set; }
        

        public VModel()
        {
            vclist = new List<VCVOUCH>();
            dclist = new List<DVOUCH>();
            mstlist = new List<MSTLST>();
            reclist = new List<REC1718>();
            vm = new List<Dashboard_New.Models.custom.vcvendor>();
            dm = new List<Dashboard_New.Models.custom.dcvendor>();
            nirf = new List<Dashboard_New.Models.custom.NIRF_RPT>();
            ppoo = new List<Dashboard_New.Models.custom.ppo>();           
        }
    }
}
