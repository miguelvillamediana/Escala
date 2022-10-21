using Escala.Models;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Collections;
using System.Diagnostics;

namespace Escala.Controllers
{
    public class HomeController : Functions
    {
        private readonly ILogger<HomeController> _logger;
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Index(IFormFile file)
        {
            emptyTempData();
            await UploadFile(file);
            return View();
        }
        
        public async Task<bool> UploadFile(IFormFile file)
        {
            string path = "";
            bool copied = false;
            string fileName = file.FileName;

            //Arrays de emails por idioma
            ArrayList ENG = new ArrayList();
            ArrayList POR = new ArrayList();
            ArrayList SPA = new ArrayList();
            ArrayList FRE = new ArrayList();
            ArrayList REPORT = new ArrayList();

            //Saving and reding file
            try
            {
                if (file.Length > 0)
                {
                    //Saving file
                    path = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "UploadedFiles"));
                    using (var fileStream = new FileStream(Path.Combine(path, fileName), FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    //Reading file
                    var newfileName = $"./UploadedFiles/{fileName}";
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    using (var stream = System.IO.File.Open(newfileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            while (reader.Read())
                            {                              
                                string EmailMissionario = "";
                                string EmailPessoal = "";
                                string IdiomaTemp = "";
                                string JustOneIdioma = "";

                                try{IdiomaTemp = reader.GetValue(15).ToString();}
                                catch{IdiomaTemp = "";}

                                //tratar los datos del idioma caso hablen mas de dos y tenga que elegir el principal
                                JustOneIdioma = VerifyIdioma(IdiomaTemp);

                                //verificar si existe el registro por idioma
                                int Spanish = JustOneIdioma.IndexOf("Spanish");
                                int Portuguese = JustOneIdioma.IndexOf("Portuguese");
                                int English = JustOneIdioma.IndexOf("English");
                                int French = JustOneIdioma.IndexOf("French");

                                //Obtener los emails 
                                try{EmailMissionario = reader.GetValue(10).ToString();}
                                catch{EmailMissionario = "";}
                                try{EmailPessoal = reader.GetValue(11).ToString();}
                                catch{EmailPessoal = "";}

                                //Adicionar cada idioma a un array
                                if (validarEmail(EmailMissionario))
                                {
                                    addToArray(SPA,POR,ENG,FRE,Spanish,English,Portuguese,French,EmailMissionario);
                                }
                                if (validarEmail(EmailPessoal))
                                {
                                    addToArray(SPA,POR,ENG,FRE,Spanish,English,Portuguese,French,EmailPessoal);
                                }
                            }
                        }
                    }

                    //Method to transfor arrays into ";" strings
                    
                    TempData["spa"] = transformArray(SPA, ";");
                    TempData["eng"] = transformArray(ENG, ";");
                    TempData["por"] = transformArray(POR, ";");
                    TempData["fre"] = transformArray(FRE, ";");

                    if (SPA.Capacity > 0) {
                        try
                        {
                            sendEmail(transformArray(SPA, ","), "SPA").GetAwaiter().GetResult();
                            REPORT.Add(SPA);
                        }
                        catch { }
                    };
                    if (ENG.Capacity > 0) {
                        try 
                        {
                            sendEmail(transformArray(ENG, ","), "ENG").GetAwaiter().GetResult();
                            REPORT.Add(ENG);
                        } 
                        catch { }
                        
                    };
                    if (POR.Capacity > 0)
                    {
                        try
                        {
                            sendEmail(transformArray(POR, ","), "POR").GetAwaiter().GetResult();
                            REPORT.Add(POR);
                        }
                        catch { }

                    };
                    if (FRE.Capacity > 0)
                    {
                        try
                        {
                            sendEmail(transformArray(FRE, ","), "FRE").GetAwaiter().GetResult();
                            REPORT.Add(FRE);
                        }
                        catch { }

                    };

                    TempData["msg"] = "Arquivo carregado com sucesso!";
                    TempData["call"] = "cartasEnviadas();";
                    copied = true;
                }
                else
                {
                    copied = false;
                }

            }
            catch(Exception e)
            {
                TempData["msg"] = "Erro ao carregar o arquivo, tente novamente";
                TempData["call"] = "erroEnvio();";
            }
            return copied;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}