using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class ppo
    {
       
        public string from_dt { get; set; }
        //[Required]
        public string to_dt { get; set; }
        //[Required]
        public string from_dt1 { get; set; }
        //[Required]
        public string to_dt1 { get; set; }
        public string indent_no { get; set; }
        public string name { get; set; }
        public System.DateTime indent_date { get; set; }      
        public string purch_desc { get; set; }
        public string category { get; set; }
        public Nullable<double> indent_value { get; set; }    
        public decimal IndentValueRs { get; set; }
        public string cocode { get; set; }
        public string gcode { get; set; }
        public string currency { get; set; }
        public string IndentType { get; set; }
        public string status { get; set; }
        public string remarks { get; set; }
        public string dept { get; set; }

    }
}