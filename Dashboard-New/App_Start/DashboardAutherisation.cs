using Dashboard_New.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
namespace Dashboard_New.App_Start
{   
    public class DashboardAutherisation : AuthorizeAttribute
    {
        FACCTEntities data = new FACCTEntities();
        private readonly string[] allowedroles;
        public DashboardAutherisation(params string[] roles)
        {
            this.allowedroles = roles;
        }
        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            bool authorize = false;
            if(HttpContext.Current.Session["username"] == null)
            {
                return false;
            }
            if (allowedroles.Length == 0)
            {
                return true;
            }  
            foreach (var role in allowedroles)
            {
                string users = (string)HttpContext.Current.Session["username"];
                //var user = data.useraccounts.Where(m => string.Equals(m.Name, users)).SingleOrDefault(); // checking active users with allowed roles.  
                string[] userroles;
                using (DashboardEntities d = new DashboardEntities())
                {                   
                    var typ = (from u in new DashboardEntities().User_Roles where u.UserName == users select u).SingleOrDefault();
                    var user_type = d.User_Roles.Where(m=>string.Equals(m.UserName,users) && string.Equals(m.UserType, typ.UserType)).FirstOrDefault();
                    userroles = user_type.reports.Split('|');
                }
                foreach(string ur in userroles)
                {
                    if (string.Equals(role, ur))
                    {
                      
                        //ViewBag.calljavascriptfunction = "Javascriptfunction();";
                        return true;
                    }
                    //else
                    //{ ViewBag.calljavascriptfunction = "Javascriptfunction();"; }
                }
                
            }
            return authorize;
        }
        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            string refurl = string.Empty;
            if (HttpContext.Current.Request.UrlReferrer != null)
            {
                 refurl = HttpContext.Current.Request.UrlReferrer.AbsoluteUri;
            }
            //if (!string.IsNullOrEmpty(refurl))
            //{
            //    string url2 = refurl.Split('?')[0];
            //    filterContext.Result = new RedirectResult(url2 + "?notauth=y");
            //}
            if (!string.IsNullOrEmpty(refurl))
            {                
                string url2 = refurl.Split('?')[0];
                string url3 = refurl.Split('/')[0];
                string url4 = refurl.Split('/')[2];
                string url5 = url3 +"//"+ url4;
                //if (refurl == url5 + "/Home/vcpost") //--old
                    if (refurl == url5 + "/dashboard/Home/vcpost")  //pub
                    {
                    //filterContext.Result = new RedirectResult("../Home/vc?notauth=y"); //old
                    filterContext.Result = new RedirectResult("../Home/vc?notauth=y");//pub                   
                }
                else
                {
                    //filterContext.Result = new RedirectResult(url2 + "?notauth=y");//old
                    filterContext.Result = new RedirectResult(url2 + "?notauth=y"); //pub
                }
            }
            else
            filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new { controller = "Login", action = "Index" }));
        }
    }
}