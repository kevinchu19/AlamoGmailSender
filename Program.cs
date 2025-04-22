using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Reflection;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using MimeKit;
using Google.Apis.Util.Store;
using System.Globalization;

class Program
{
    static async Task Main(string[] args)
    {
        if (args.Length < 8)
        {
            Console.WriteLine("Uso: programa <origen> <destinos> <cc> <cco> <archivosAdjuntos> <asunto> <cuerpoHtml> <archivoLog>");
            return;
        }

        string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        string origenRaw = args[0];
        string destinosRaw = args[1];
        string ccRaw = args[2];
        string ccoRaw = args[3];
        string archivosAdjuntosRaw = args[4];
        string asunto = args[5];
        string cuerpoHtml = args[6];
        string archivoLog = args[7];

        string emailOrigen, nombreOrigen = "", dominio;

        // Intentar matchear formato "Nombre <correo@dominio>"
        var match = System.Text.RegularExpressions.Regex.Match(origenRaw, @"^(.*)\s*<(.+@(.+))>$");

        if (match.Success)
        {
            nombreOrigen = match.Groups[1].Value.Trim();
            emailOrigen = match.Groups[2].Value.Trim();
            dominio = match.Groups[3].Value.Trim().ToLower();
        }
        else
        {
            // Formato simple: correo@dominio
            emailOrigen = origenRaw.Trim();
            var emailParts = emailOrigen.Split('@');
            dominio = emailParts.Length == 2 ? emailParts[1].ToLower() : "default";
        }

        // Ruta esperada del archivo de credenciales
        string archivoCredenciales = Path.Combine(exeDirectory, $"credentials_{dominio}.json");

        // Ruta para token por mail de origen
        string tokenDirectory = Path.Combine(exeDirectory, $"token_{emailOrigen}");

        try
        {
            Log(archivoLog, "Inicio del proceso de envío de correo.");

            if (!File.Exists(archivoCredenciales))
            {
                Log(archivoLog, $"No se encontró el archivo: {archivoCredenciales}");
                return;
            }

            Log(archivoLog, $"Autenticando con OAuth2 usando: {archivoCredenciales}");
            UserCredential credential;
            try
            {
                using (var stream = new FileStream(archivoCredenciales, FileMode.Open, FileAccess.Read))
                {
                    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        new[] { GmailService.Scope.MailGoogleCom },
                        "user",
                        CancellationToken.None,
                        new FileDataStore(tokenDirectory, true)
                    );
                }
            }
            catch (Exception authEx)
            {
                Log(archivoLog, $"ERROR en autenticación: {authEx.Message}");
                if (Directory.Exists(tokenDirectory))
                {
                    Log(archivoLog, $"Eliminando directorio de token corrupto: {tokenDirectory}");
                    Directory.Delete(tokenDirectory, true);
                }
                return;
            }

            Log(archivoLog, "Autenticación exitosa.");

            var service = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail API C#"
            });

            var message = new MimeMessage();
            if (!string.IsNullOrWhiteSpace(nombreOrigen))
                message.From.Add(new MailboxAddress(nombreOrigen, emailOrigen));
            else
                message.From.Add(new MailboxAddress(emailOrigen, emailOrigen));

            var destinosList = destinosRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
            foreach (var destino in destinosList)
                message.To.Add(new MailboxAddress("", destino));
            Log(archivoLog, $"Destinatarios agregados: {string.Join(", ", destinosList)}");

            if (!string.IsNullOrWhiteSpace(ccRaw))
            {
                var ccList = ccRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
                foreach (var cc in ccList)
                    message.Cc.Add(new MailboxAddress("", cc));
                Log(archivoLog, $"CC agregado: {string.Join(", ", ccList)}");
            }

            if (!string.IsNullOrWhiteSpace(ccoRaw))
            {
                var ccoList = ccoRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
                foreach (var cco in ccoList)
                    message.Bcc.Add(new MailboxAddress("", cco));
                Log(archivoLog, $"CCO agregado: {string.Join(", ", ccoList)}");
            }

            message.Subject = asunto;
            var bodyBuilder = new BodyBuilder { HtmlBody = cuerpoHtml };

            if (!string.IsNullOrWhiteSpace(archivosAdjuntosRaw))
            {
                var archivosAdjuntos = archivosAdjuntosRaw.Split(';').Select(a => a.Trim()).Where(a => !string.IsNullOrWhiteSpace(a)).ToList();
                foreach (var archivo in archivosAdjuntos)
                {
                    if (File.Exists(archivo))
                    {
                        bodyBuilder.Attachments.Add(archivo);
                        Log(archivoLog, $"Archivo adjunto agregado: {archivo}");
                    }
                    else
                    {
                        Log(archivoLog, $"ADVERTENCIA: El archivo no existe y no será adjuntado: {archivo}");
                    }
                }
            }
            else
            {
                Log(archivoLog, "No se adjuntarán archivos.");
            }

            message.Body = bodyBuilder.ToMessageBody();

            using (var memoryStream = new MemoryStream())
            {
                message.WriteTo(memoryStream);
                var rawMessage = Convert.ToBase64String(memoryStream.ToArray())
                    .Replace("+", "-")
                    .Replace("/", "_")
                    .Replace("=", "");

                var gmailMessage = new Message { Raw = rawMessage };
                await service.Users.Messages.Send(gmailMessage, "me").ExecuteAsync();
            }

            Log(archivoLog, "Correo enviado exitosamente.");
        }
        catch (Exception ex)
        {
            Log(archivoLog, $"Error al enviar el correo: {ex.Message}");
        }
    }

    static void Log(string archivoLog, string mensaje)
    {
        try
        {
            using (StreamWriter writer = new StreamWriter(archivoLog, true))
            {
                string timestamp = DateTime.Now.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
                writer.WriteLine($"{timestamp} - AlamoGmailSender - {mensaje}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al escribir en el archivo de log: {ex.Message}");
        }
    }
}
