using Escala.Models;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace Escala.Controllers
{
    public class HomeController : Controller
    {
        //Clase para mantener un registro e log en caso de errores
        private readonly ILogger<HomeController> _logger;
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        //Vista del Index
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Index(IFormFile file)
        {
            //vaciar tempdatas
            TempData["eng"] = "";
            TempData["por"] = "";
            TempData["spa"] = "";
            TempData["fre"] = "";
            await UploadFile(file);
            return View();
        }

        //Method to upload and catch a file
        public async Task<bool> UploadFile(IFormFile file)
        {
            List<NewMissionaries> listaMissionarios = new List<NewMissionaries>();
            string path = "";
            bool copied = false;
            string fileName = file.FileName;

            //Arrays de emails por idioma
            ArrayList ENG = new ArrayList();
            ArrayList POR = new ArrayList();
            ArrayList SPA = new ArrayList();
            ArrayList FRE = new ArrayList();

            try
            {
                if (file.Length > 0)
                {
                    path = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "UploadedFiles"));
                    using (var fileStream = new FileStream(Path.Combine(path, fileName), FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    var newfileName = $"./UploadedFiles/{fileName}";

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    using (var stream = System.IO.File.Open(newfileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {

                            while (reader.Read()) //Each row of the file
                            {
                                //Variables para separar los idiomas de cada row
                                string IdiomaTemp = reader.GetValue(15).ToString();

                                //tratar los datos del idioma caso hablen mas de dos y tenga que elegir el principal
                                string JustOneIdioma = VerifyIdioma(IdiomaTemp);

                                //verificar si existe el registro por idioma
                                int Spanish = JustOneIdioma.IndexOf("Spanish");
                                int Portuguese = JustOneIdioma.IndexOf("Portuguese");
                                int English = JustOneIdioma.IndexOf("English");
                                int French = JustOneIdioma.IndexOf("French");

                                //Obtener los emails 
                                string EmailMissionario = reader.GetValue(10).ToString();
                                string EmailPessoal = reader.GetValue(11).ToString();

                                //Adicionar cada idioma a un array
                                if (Spanish == 0)
                                {
                                    if (validarEmail(EmailMissionario))
                                    {
                                        SPA.Add(EmailMissionario);
                                    }
                                    if (validarEmail(EmailPessoal))
                                    {
                                        SPA.Add(EmailPessoal);
                                    }                                    
                                    
                                }
                                if (English == 0)
                                {
                                    if (validarEmail(EmailMissionario))
                                    {
                                        ENG.Add(EmailMissionario);
                                    }
                                    if (validarEmail(EmailPessoal))
                                    {
                                        ENG.Add(EmailPessoal);
                                    }

                                }
                                if (Portuguese == 0)
                                {
                                    if (validarEmail(EmailMissionario))
                                    {
                                        POR.Add(EmailMissionario);
                                    }
                                    if (validarEmail(EmailPessoal))
                                    {
                                        POR.Add(EmailPessoal);
                                    }

                                }
                                if (French == 0)
                                {
                                    if (validarEmail(EmailMissionario))
                                    {
                                        FRE.Add(EmailMissionario);
                                    }
                                    if (validarEmail(EmailPessoal))
                                    {
                                        FRE.Add(EmailPessoal);
                                    }
                                }
                                /*
                                listaMissionarios.Add(new NewMissionaries
                                {
                                    Idioma = reader.GetValue(15).ToString(),
                                    EmailMissionario = reader.GetValue(10).ToString(),
                                    EmailPessoal = reader.GetValue(11).ToString()
                                });
                                */
                            }
                        }
                    }

                    //Method to transfor arrays into ";" strings
                    

                    TempData["spa"] = transformArray(SPA);
                    TempData["eng"] = transformArray(ENG);
                    TempData["por"] = transformArray(POR);
                    TempData["fre"] = transformArray(FRE);
                    

                    TempData["msg"] = "File uploaded successfully";
                    copied = true;
                }
                else
                {
                    copied = false;
                }

            }
            catch
            {
                TempData["msg"] = "Erro ao carregar o arquivo, tente novamente";
            }
            return copied;
        }

        // Funcion para tratar los idiomas caso ellos hablen mas de uno
        public string VerifyIdioma(string idioma)
        {
            try
            {
                string[] listaIdiomas = idioma.Split(";");
                idioma = listaIdiomas[0];
            }
            catch
            {}
            return idioma;
        }

        public string transformArray(ArrayList emails)
        {
            string[] StringSArr = (string[])emails.ToArray(typeof(string));
            var str = String.Join(";", StringSArr);

            return str;
        }
            public static bool validarEmail(string email)
            {
                if (string.IsNullOrWhiteSpace(email))
                    return false;

                try
                {
                    // Normalize the domain
                    email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                                          RegexOptions.None, TimeSpan.FromMilliseconds(200));

                    // Examines the domain part of the email and normalizes it.
                    string DomainMapper(Match match)
                    {
                        // Use IdnMapping class to convert Unicode domain names.
                        var idn = new IdnMapping();

                        // Pull out and process domain name (throws ArgumentException on invalid)
                        string domainName = idn.GetAscii(match.Groups[2].Value);

                        return match.Groups[1].Value + domainName;
                    }
                }
                catch (RegexMatchTimeoutException e)
                {
                    return false;
                }
                catch (ArgumentException e)
                {
                    return false;
                }

                try
                {
                    return Regex.IsMatch(email,
                        @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                        RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
                }
                catch (RegexMatchTimeoutException)
                {
                    return false;
                }
            }

        public void sendEmail()
        {
            MailMessage mail = new MailMessage();
            mail.To.Add("jose.villamediana.osorio@gmail.com");
            mail.From = new MailAddress("noreplyctmbrasil@gmail.com");
            mail.Subject = "Teste email - Programa escala :)";
            mail.Body = "<p>Se esse teste chegou, todo é possivel<br/>. Ate 'What up G'.</p>";
            mail.IsBodyHtml = true;

            SmtpClient smtp = new SmtpClient();
            smtp.Port = 587; // 25 465
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Host = "smtp.gmail.com";
            smtp.Credentials = new System.Net.NetworkCredential("noreplyctmbrasil@gmail.com", "minvtnwbdfzxlfbi");
            smtp.Send(mail);
        }

    //Parametros para retornar la vista del error        
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}