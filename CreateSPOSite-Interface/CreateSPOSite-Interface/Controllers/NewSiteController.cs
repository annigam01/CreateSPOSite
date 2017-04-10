using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CreateSPOSite_Interface.Controllers
{
    public class NewSiteController : Controller
    {
        // GET: NewSite
        public string NewSite()
        {
            return "NewSite page invoked!";
        }

        public ActionResult GetView()
        {
            return View("NewSiteView");

        }
    }
}