using Microsoft.AspNetCore.Mvc;
using System.Collections;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace Escala.Controllers
{
    public class Functions : Controller
    {
        public void emptyTempData()
        {
            TempData["eng"] = "";
            TempData["por"] = "";
            TempData["spa"] = "";
            TempData["fre"] = "";
        }
        public void addToArray(ArrayList SPA, ArrayList POR, ArrayList ENG, ArrayList FRE, int spa, int eng, int por, int fre, string email)
        {
            if (spa == 0) { SPA.Add(email); }
            if (eng == 0) { POR.Add(email); }
            if (por == 0) { ENG.Add(email); }
            if (fre == 0) { FRE.Add(email); }

        }
        public string transformArray(ArrayList emails, string separador)
        {
            string[] StringSArr = (string[])emails.ToArray(typeof(string));
            var str = String.Join(separador, StringSArr);

            return str;
        }
        public string VerifyIdioma(string idioma)
        {
            try
            {
                string[] listaIdiomas = idioma.Split(";");
                idioma = listaIdiomas[0];
            }
            catch
            { }
            return idioma;
        }
        public static bool validarEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;
            try
            {
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,RegexOptions.None, TimeSpan.FromMilliseconds(200));
                string DomainMapper(Match match)
                {
                    var idn = new IdnMapping();
                    string domainName = idn.GetAscii(match.Groups[2].Value);
                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException e)
            {return false;}
            catch (ArgumentException e)
            {return false;}

            try
            {
                return Regex.IsMatch(email,
                    @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {return false;}
        }
        public async Task sendEmail(string emails, string idioma)
        {
            string emailMaster = "";
            string password = "";
            string emailCopy = "";
            try
            {
                SmtpClient smtp = new SmtpClient();
                smtp.Port = 587; // 25 465
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Host = "smtp.gmail.com";
                smtp.Credentials = new NetworkCredential(emailMaster, password);

                MailMessage mail = new MailMessage();
                mail.To.Add(emails);
                mail.From = new MailAddress(emailMaster);
                MailAddress copy = new MailAddress(emailCopy);
                mail.CC.Add(copy);

                //variables por idioma
                string title = "";
                string supervisor = "Bueno";
                string subject = "";
                string gretting = "";
                string message1 = "";
                string message2 = "";
                string message3 = "";
                string link = "";
                string bye = "";
                string bye2 = "";

                if (idioma == "SPA")
                {
                    subject = "BIENVENIDO AL CCM ONLINE!";
                    gretting = "Hola";
                    message1 = "¡Felicidades por su llamamiento misional! ¡Estamos muy felices por su decisión de servir al Señor en Su gran obra!";
                    message2 = $"Me llamo Hermana {supervisor} y soy responsable por el Pré-CCM aquí en Brasil. Sé que usted necesita aprender un nuevo idioma en su misión. El Pré-CCM fue creado para ayudarle a comenzar a aprender el idioma de su misión ¡incluso antes de entrar en el CCM!";
                    message3 = "Por favor, llene el formulario añadido a este correo electrónico para que podamos coordinar sus clases con nuestros instructores. Cuando su instructor se comunique con usted, determinarán juntos qué hora dentro del período elegido funciona mejor para usted, junto con la duración de la clase, el día y la cantidad de clases por semana. <mark>Las clases pueden comenzar a partir del próximo Lunes.</mark>";
                    link = "https://forms.gle/2VTaRt2AKrt3R45K6";
                    bye = "¡Puede hacer cualquier otra pregunta que pueda tener cuando usted y su instructor se reúnan por primera vez!";
                    bye2 = "Tenga una linda semana!";
                    title = "Hermana";
                }
                if (idioma == "POR")
                {

                    subject = "BEM-VINDO AO CTM ONLINE!";
                    gretting = "Olá";
                    message1 = "Parabéns pelo seu chamado missionário! Estamos muito felizes pela sua decisão de servir ao Senhor em Sua grande obra!";
                    message2 = $"Me chamo Irmã {supervisor} e sou responsável pelo Pré-CTM aqui no Brasil.Sei que você precisa aprender um novo idioma para a sua missão. O Pré-CTM foi criado para ajudar você a começar a aprender o idioma da sua missão antes mesmo de entrar no CTM!";
                    message3 = "Por favor, preencha o formulário anexado neste e-mail para que possamos coordenar suas aulas com os nossos instrutores. Quando seu instrutor te contatar vocês combinarão juntos qual horário dentro do período escolhido funciona melhor para vocês, juntamente com a duração da aula, dia e quantidade de aulas na semana. <mark>As aulas podem começar a partir da próxima segunda-feira.</mark>";
                    link = "https://forms.gle/yZBK8GEcFtjX117p9";
                    bye = "Você pode fazer quaisquer outras perguntas que possa ter quando você e seu instrutor se reunirem pela primeira vez!";
                    bye2 = "Tenha uma ótima semana!";
                    title = "Irmã";
                }
                if (idioma == "ENG")
                {
                    subject = "WELCOME TO ONLINE MTC!";
                    gretting = "Hello";
                    message1 = "Congratulations on your mission call! We are very happy for your decision to serve the Lord in His great work!";
                    message2 = $"My name is Sister {supervisor} and I am responsible for the Pre-MTC program here in Brazil. I know you need to learn a new language for your mission. Pre-MTC is designed to help you start learning your mission language before you even enter the MTC!";
                    message3 = "Please fill out the form attached to this email so that we can coordinate your classes with our instructors.When your instructor contacts you, you will work out together what time within the chosen period works best for you, along with the class duration, day and number of classes per week. <mark>Classes may start from next Monday on.</mark>";
                    link = "https://forms.gle/rP8deFj9ge9ntvGQ7";
                    bye = "You can ask any other questions you may have when you and your instructor meet for the first time!";
                    bye2 = "Have a great week!";
                    title = "Sister";
                }
                if (idioma == "FRE")
                {
                    subject = "";
                    gretting = "";
                    message1 = "";
                    message2 = "";
                    message3 = "";
                    link = "";
                    bye = "";
                    bye2 = "";
                }

                mail.Subject = $"{subject}";
                mail.Body = $"<p>{gretting},</p>" +
                    $"<p>{message1}</p>" +
                    $"<p>{message2}</p>" +
                    $"<p>{message3}</p>" +
                    $"<a href='{link}' style='color: blue; text-style: underline;'>{link}<a/>" +
                    $"<p>{bye}</p>" +
                    $"<p>{bye2}</p>" +
                    $"<br><p>{title} {supervisor}.</p>";

                mail.IsBodyHtml = true;
                using (smtp)
                {
                    await smtp.SendMailAsync(mail);
                }
            }
            catch (Exception e)
            {
                TempData["msg"] = "Erro ao carregar o arquivo, tente novamente";
                TempData["call"] = "erroEnvio();";
                throw;
            }
        }
    }
}
