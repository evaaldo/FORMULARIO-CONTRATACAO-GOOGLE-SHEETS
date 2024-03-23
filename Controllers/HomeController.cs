using FormularioGoogleSheets.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace FormularioGoogleSheets.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private readonly string _credentialsFilePath = @"C:\Chaves API\credentials.json";
        private readonly string _spreadsheetId = "1JM88RPurmCArPrBDAYfsY1-j2Pq0J6msjdKZm-Qkygg";
        private readonly string _range = "Pagina1!A2:D";

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Index(Contratado dados)
        {
            try
            {
                // Autenticação usando credenciais de serviço
                GoogleCredential credential;

                using (var stream = new FileStream(_credentialsFilePath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }

                // Criar a service do Google Sheets
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Integracao Google Sheets",
                });

                // Obter dados atuais da planilha
                SpreadsheetsResource.ValuesResource.GetRequest getRequest = service.Spreadsheets.Values.Get(_spreadsheetId, _range);
                ValueRange currentValue = getRequest.Execute();

                // Determinar a próxima linha em branco
                int nextRow = currentValue.Values != null ? currentValue.Values.Count + 1 : 2;

                // Criar e enviar os dados
                ValueRange valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>>
                {
                    new List<object> { dados.NomeContratado, dados.CPF, dados.Setor, dados.Cargo },
                };

                string appendRange = $"Pagina1!A{nextRow}:D{nextRow}";

                SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, appendRange);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

                AppendValuesResponse appendResponse = await appendRequest.ExecuteAsync();

                return RedirectToAction("Index");
            }
            catch (Google.GoogleApiException apiException)
            {
                return BadRequest(apiException.Error.Code);
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
