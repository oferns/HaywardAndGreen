using System.Web.Mvc;

namespace EmailMonitor.Controllers
{
    public class EmailMonitorController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }


        [HttpGet]
        public ActionResult ServerInfor()
        {
            return View();
        }
    }
}