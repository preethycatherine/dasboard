using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Dashboard_New.Models;

namespace Dashboard_New.Controllers
{
    public class UserController : Controller
    {
        public ActionResult Index()
        {
            List<UserModel> users = UserModel.getUsers();
            return View(users);
        }

        public JsonResult UpdateUser(UserModel model)
        {
            // Update model to your db
            string message = "Success";
            return Json(message, JsonRequestBehavior.AllowGet);
        }

    }
}
