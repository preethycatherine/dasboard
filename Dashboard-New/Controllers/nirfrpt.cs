using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Dashboard_New.Models;

namespace Dashboard_New.Controllers
{
    public class Home : Controller
    {
        // GET: nirfrpt
        public ActionResult nirf()
        {
            vcEntities obj = new vcEntities();
         
            List<REC1718> rec = obj.REC1718.ToList();
            List<MSTLST> mst = obj.MSTLSTs.ToList();
            return View();
        }
    }
}