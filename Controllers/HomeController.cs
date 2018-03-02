using System.IO;
using System.Web;
using System.Web.Mvc;

namespace ImportToXLS.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {            
            var tamanhoConteudo = Request.Files["file"].ContentLength;
            var nomeArquivo = Request.Files["file"].FileName;
            var extensaoArquivo = Path.GetExtension(nomeArquivo);

            if (extensaoArquivo == ".xls" || extensaoArquivo == ".xlsx")
            {
                ImportarCpfPlanilhaExcelCsv.TratarArquivoExcel(tamanhoConteudo, extensaoArquivo, nomeArquivo, Request.Files["file"].InputStream);                
            }
            else
            {
                ImportarCpfPlanilhaExcelCsv.TratarArquivoCsv(nomeArquivo, Request.Files["file"].InputStream);
            }           

            return View();
        }       

    }
}
