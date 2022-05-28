using Dapper;
using Dapper.Contrib.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;
using TableAttribute = Dapper.Contrib.Extensions.TableAttribute;
using DocumentFormat.OpenXml.CustomXmlSchemaReferences;
using DocumentFormat.OpenXml.Spreadsheet;
using MySqlX.XDevAPI;

namespace ExcelToDataBase.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }
        [Consumes("multipart/form-data")]
        [HttpPost("input-file")]
        public ActionResult InputFile(IFormFile file)
        {
            var streamFile = ReadStream(file);


            var disciplina = ReadXls(streamFile);

            SaveDisciplina(disciplina);

            return Ok();
        }


        private static List<Disciplina> ReadXls(MemoryStream stream)
        {
            var response = new List<Disciplina>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCoumt = worksheet.Dimension.End.Column;

                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    var disciplina = new Disciplina();
                    disciplina.Nome = worksheet.Cells[row, 1].Value.ToString();
                    disciplina.Id = Convert.ToInt32(worksheet.Cells[row, 2].Value.ToString());
                    disciplina.periodo = Convert.ToInt32(worksheet.Cells[row, 3].Value.ToString());
                    disciplina.categoria = worksheet.Cells[row, 4].Value.ToString();
                    disciplina.dificuldade = Convert.ToDecimal(worksheet.Cells[row, 5].Value.ToString());
                    disciplina.Creditos = Convert.ToInt32(worksheet.Cells[row, 6].Value.ToString());
                    disciplina.HoraAula = Convert.ToInt32(worksheet.Cells[row, 7].Value.ToString());
                    disciplina.HoraRelogio = Convert.ToInt32(worksheet.Cells[row, 8].Value.ToString());
                    disciplina.QtdTeorica = Convert.ToInt32(worksheet.Cells[row, 9].Value.ToString());
                    disciplina.QtdPratica = Convert.ToInt16(worksheet.Cells[row, 10].Value.ToString());
                    disciplina.Ementa = worksheet.Cells[row, 11].Value.ToString();

                    response.Add(disciplina);
                }

            }
            return response;

        }

    private static void SaveDisciplina(List<Disciplina> disciplina)
    {
             using (var connection = new SqlConnection(connectionString: "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=Curso;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"))
                //Teste string
            //using (var connection = new SqlConnection(connectionString: "Server=(localdb)\\MSSQLLocalDB;Database=Curso;Trusted_Connection=true;"))
            {
                connection.Insert(disciplina);
        }
    }


    [Table("Disciplina")]
  
    public class Disciplina
    {        
        public string Nome { get; set; }
            [ExplicitKey]
            public int Id { get; set; }
        public int periodo { get; set; }
        public string categoria { get; set; }
        public decimal dificuldade { get; set; }
        public int Creditos { get; set; }
        public int HoraAula { get; set; }
        public int HoraRelogio { get; set; }
        public int QtdTeorica { get; set; }
        public int QtdPratica { get; set; }
        public string Ementa { get; set; }
    }

    protected  MemoryStream ReadStream(IFormFile formFile)
    {
        using (var stream = new MemoryStream())
          {
            formFile?.CopyTo(stream);

            var byteArray = stream.ToArray();

            return new MemoryStream(byteArray);
         }
      }

   }
}
