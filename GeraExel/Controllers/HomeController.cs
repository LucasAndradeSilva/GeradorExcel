using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using GeraExel.Models;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.AspNetCore.Http;
using System.Text;

namespace GeraExel.Controllers
{
   
    public class model
    {
        public string name { get; set; }
        public string[] data { get; set; }
    }
    public class HomeController : Controller
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

        public T[] GetObjectArray<T>(string json, string mainTag) where T : new()
        {
            T[] instance = null;
            try
            {
                instance = JToken.Parse(json)[mainTag].ToObject<T[]>();
            }
            catch (Exception e)
            {
                var teste = e;
            }
            return instance;
        }


        [Route("Guarda")]
        public IActionResult Guarda(string json)
        {
            try
            {
                HttpContext.Session.Set("DataChart", Encoding.UTF8.GetBytes(json));
                return Json("");
            }
            catch (Exception ex)
            {

                return Json(ex.Message);
            }
        }
        [Route("Excel")]
        public IActionResult GerarExcel()
        {
            try
            {
                var json = Encoding.UTF8.GetString(HttpContext.Session.Get("DataChart"));
                var dadosGrafico = GetObjectArray<model>(json, "dados");

                DataSet ds = new DataSet("New_DataSet");

                DataTable dt = new DataTable("Teste1");

                dt.Columns.Add("Nome");
                dt.Columns.Add("Data");        
                for (int i = 0; i < dadosGrafico.Length; i++)
                {
                    dt.Rows.Add(dadosGrafico[i].name, dadosGrafico[i].data);
                }

                ds.Tables.Add(dt);

                MemoryStream stream = new MemoryStream();
                ExcelLibrary.DataSetHelper.CreateWorkbook(stream, ds);
                
                var contentType = "application/vnd.ms-excel";
                var handle = ("content-disposition", string.Format("attachment;filename=Teste_{0}.xls", DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")));

                stream.WriteTo(stream);
                
                return File(stream.ToArray(), contentType, "Teste.xls");
            }
            catch (Exception ex)
            {
               return Json("Erro : " + ex.Message);
            }         
        }

        private void liberarObjetos(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;               
            }
            finally
            {
                GC.Collect();
            }
        }

        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-Excel", fileName);
            }
            else
            {
                // Problem - Log the error, generate a blank file,
                //           redirect to another controller action - whatever fits with your application
                return new EmptyResult();
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
