using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models.custom
{
    public class chqmdf
    {
        public string from_dt { get; set; }
        public string to_dt { get; set; }
        public List<chqdrawn> chqdrawn_dup { get; set; }
        public chqmdf()
        {
            chqdrawn_dup = new List<chqdrawn>();
        }
    }
}