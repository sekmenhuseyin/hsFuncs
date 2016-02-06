using System.Web.Mvc;
using hsFunctions;

namespace TesterMvc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.encrypted = functions.Encrypt("Hüseyin Sekmenoğlu");
            ViewBag.decrypted = functions.Decrypt(ViewBag.encrypted);
            return View();
        }
    }
}