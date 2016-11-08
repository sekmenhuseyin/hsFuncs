using System.Web.Mvc;
using hsFunctions;

namespace TesterMvc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.encrypted = fns.EncryptString("Hüseyin Sekmenoğlu","KarineSoft");
            ViewBag.decrypted = fns.DecryptString(ViewBag.encrypted, "KarineSoft");
            return View();
        }
    }
}