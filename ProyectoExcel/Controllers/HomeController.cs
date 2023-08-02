using Microsoft.AspNetCore.Mvc;
using ProyectoExcel.Models;
using System.Diagnostics;

using NPOI.SS.UserModel;    //Usar Modelo
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;  //Leer solo un archivo (xlsx)
using ProyectoExcel.Models.ViewModels;
using EFCore.BulkExtensions;

namespace ProyectoExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly DbExcelContext _dbcontext;

        public HomeController(DbExcelContext context)
        {
            _dbcontext = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult MostrarDatos([FromForm] IFormFile ArchivoExcel)
        {
            Stream stream = ArchivoExcel.OpenReadStream();   //Leer el archivo excel en memoria

            IWorkbook MiExcel = null;

            if(Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream); //El archivo sea leidoo con este recurso
            }
            else
            {
                MiExcel = new XSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0); //Obtener solo la primera Hoja

            int cantidadFilas = HojaExcel.LastRowNum;    //Cantidad de filas de la hoja excel

            List<VMContacto> Lista = new List<VMContacto>();

            for(int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);    //Devolver una hoja por cada iteracion

                Lista.Add(new VMContacto
                {
                    nombre = fila.GetCell(0).ToString(),
                    apellido = fila.GetCell(1).ToString(),
                    telefono = fila.GetCell(2).ToString(),
                    correo = fila.GetCell(3).ToString()
                });
            }

            return StatusCode(StatusCodes.Status200OK, Lista); //Devuelve en forma de lista
        }

        //Envio de Datos
        [HttpPost]
        public IActionResult EnviarDatos([FromForm] IFormFile ArchivoExcel)
        {
            Stream stream = ArchivoExcel.OpenReadStream();   //Leer el archivo excel en memoria

            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream); //El archivo sea leidoo con este recurso
            }
            else
            {
                MiExcel = new XSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0); //Obtener solo la primera Hoja

            int cantidadFilas = HojaExcel.LastRowNum;    //Cantidad de filas de la hoja excel

            List<Contacto> Lista = new List<Contacto>();

            for (int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);    //Devolver una hoja por cada iteracion

                Lista.Add(new Contacto
                {
                    Nombre = fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Correo = fila.GetCell(3).ToString()
                });
            }

            _dbcontext.BulkInsert(Lista);

            return StatusCode(StatusCodes.Status200OK, new { mensaje = "ok"});  //Devuelve en forma de lista
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