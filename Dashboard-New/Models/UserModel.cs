using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Dashboard_New.Models
{
    public class UserModel
    {       
        public string PARTY { get; set; }
        public string CHEQ_NO { get; set; }

        public static List<UserModel> getUsers()
        {
            List<UserModel> users = new List<UserModel>()
                {
                     new UserModel (){ PARTY="Jon", CHEQ_NO="3455" },
                     new UserModel (){  PARTY="Alex", CHEQ_NO="345345" }
                };
            return users;
        }
    }
}