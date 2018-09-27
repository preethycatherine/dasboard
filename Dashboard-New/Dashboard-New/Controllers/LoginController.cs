using Dashboard_New.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Data;

namespace Dashboard_New.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }
       
         public ActionResult Login_page()
        {
            return View();
        }
     
        [HttpPost]
        public ActionResult Index(string username, string password)
        {
            using (FACCTEntities lge = new FACCTEntities())
            {
               var v=lge.useraccounts.Where(m=>m.UserName.Trim()==username.Trim() && m.Password==password).ToList();
                //var v = lge.useraccounts.Where(a => string.Equals(a.Name,username) && string.Equals(a.Password,password)).FirstOrDefault();
                if (v !=null)
                {
                    Session["role"] = "admin";
                    Session["username"] = username;
                    Session["pass"] = password;
                   Session["isUservalid"] = true;
                    Session["loginExpiry"] = DateTime.Now.AddMinutes(2);
                    if (authenticateuser())
                    {
                        return RedirectToAction("Index", "Home", new { });
                    }
                }
                ViewData["Message"] = "u";
                return View("Index");

                // return View("Alert"); 

            }
        }
        public  bool authenticateuser()
        {
            using (DashboardEntities d = new DashboardEntities())
            {
                string user = (string)Session["username"]??"";
                string pass = (string)Session["pass"]??"";
               account_access ac = new account_access();
                try
                {
                    ac = d.account_access.Where(a => string.Equals(a.UserName, user) && string.Equals(a.Password, pass)).FirstOrDefault();
                }
                catch (Exception ex)
                {
                }
                if (ac != null)
                {
                    return true;
                }              
                 
               
                return false;
            }
        }
        public ActionResult LogOut()
        {
            //FormsAuthentication.SignOut();
            Session["role"] = "";
            Session["username"] = null;
            Session["pass"] = null;
            Session["isUservalid"] = false;
            return RedirectToAction("Index", "Login");
        }
        //[HttpPost]
        //public ActionResult dashboard()
        //{
        //    using (DashboardEntities d = new DashboardEntities())
        //    {
        //        var v = d.account_access.Where(a => string.Equals(a.UserName, Session["username"]) && string.Equals(a.Password, Session["pass"])).FirstOrDefault();
        //        if (v != null)
        //        {
        //            return RedirectToAction("Index", "Home", new { });
        //        }
        //            return View("Index");
        //    }
        //}

    }
}