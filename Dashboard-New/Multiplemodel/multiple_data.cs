using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Dashboard_New.Models;
namespace Dashboard_New.Multiplemodel
{
    public class multiple_data
    {
        public IEnumerable<REC1718> recs { get; set; }
        public IEnumerable<MSTLST> msts { get; set; }

      
    }
}