using System.Web.Mvc;
using hsFunctions;

namespace TesterMvc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.encrypted = fn.Encrypt2("Hüseyin Sekmenoğlu");
            ViewBag.decrypted = fn.Decrypt2(ViewBag.encrypted);
            return View();
        }
    }
}