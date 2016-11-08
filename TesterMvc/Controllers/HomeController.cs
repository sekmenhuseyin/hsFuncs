using System.Web.Mvc;
using hsFunctions;

namespace TesterMvc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.encrypted = fns.Encrypt("Hüseyin Sekmenoğlu","KarineSoft");
            ViewBag.decrypted = fns.Decrypt(ViewBag.encrypted, "KarineSoft");
            return View();
        }
    }
}