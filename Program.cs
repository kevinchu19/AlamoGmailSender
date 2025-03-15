using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Reflection; // Para obtener la ruta del ejecutable
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

        // Obtener la carpeta donde se encuentra el ejecutable (.exe)
        string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        // Definir rutas absolutas para credentials.json y token.json en la misma carpeta que el ejecutable
        string credentialPath = Path.Combine(exeDirectory, "credentials.json");
        string tokenPath = Path.Combine(exeDirectory, "token.json");

        string origenRaw = args[0];
        string destinosRaw = args[1];
        string ccRaw = args[2];
        string ccoRaw = args[3];
        string archivosAdjuntosRaw = args[4];
        string asunto = args[5];
        string cuerpoHtml = args[6];
        string archivoLog = args[7];

        try
        {
            Log(archivoLog, "Inicio del proceso de envío de correo.");

            if (!File.Exists(credentialPath))
            {
                Log(archivoLog, $"No se encontró el archivo credentials.json en la ruta: {credentialPath}");
            }

            Log(archivoLog, "Autenticando con OAuth2.");
            UserCredential credential;
            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] { GmailService.Scope.MailGoogleCom },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenPath, true) // Guarda el token en la misma carpeta que el .exe
                );
            }

            Log(archivoLog, "Autenticación exitosa.");

            // Crear el servicio de Gmail
            var service = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail API C#"
            });

            // **PARSEAR EL ORIGEN** (Soporta "Nombre <correo@dominio.com>")
            var match = System.Text.RegularExpressions.Regex.Match(origenRaw, @"^(.*)\s*<(.+@.+)>$");
            string nombreOrigen = match.Success ? match.Groups[1].Value.Trim() : "";
            string emailOrigen = match.Success ? match.Groups[2].Value.Trim() : origenRaw;

            var message = new MimeMessage();
            if (!string.IsNullOrWhiteSpace(nombreOrigen))
                message.From.Add(new MailboxAddress(nombreOrigen, emailOrigen));
            else
                message.From.Add(new MailboxAddress(emailOrigen, emailOrigen));

            // **AGREGAR DESTINATARIOS**
            var destinosList = destinosRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
            foreach (var destino in destinosList)
                message.To.Add(new MailboxAddress("", destino));

            Log(archivoLog, $"Destinatarios agregados: {string.Join(", ", destinosList)}");

            // **AGREGAR CC**
            if (!string.IsNullOrWhiteSpace(ccRaw))
            {
                var ccList = ccRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
                foreach (var cc in ccList)
                    message.Cc.Add(new MailboxAddress("", cc));

                Log(archivoLog, $"CC agregado: {string.Join(", ", ccList)}");
            }

            // **AGREGAR CCO**
            if (!string.IsNullOrWhiteSpace(ccoRaw))
            {
                var ccoList = ccoRaw.Split(';').Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();
                foreach (var cco in ccoList)
                    message.Bcc.Add(new MailboxAddress("", cco));

                Log(archivoLog, $"CCO agregado: {string.Join(", ", ccoList)}");
            }

            message.Subject = asunto;

            // **CREAR EL CUERPO DEL MENSAJE (HTML)**
            var bodyBuilder = new BodyBuilder { HtmlBody = cuerpoHtml };

            // **AGREGAR ARCHIVOS ADJUNTOS**
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

            // **CONVERTIR EL MENSAJE A BASE64 URL-SAFE**
            using (var memoryStream = new MemoryStream())
            {
                message.WriteTo(memoryStream);
                var rawMessage = Convert.ToBase64String(memoryStream.ToArray())
                    .Replace("+", "-")
                    .Replace("/", "_")
                    .Replace("=", "");

                // **ENVIAR EL CORREO**
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
