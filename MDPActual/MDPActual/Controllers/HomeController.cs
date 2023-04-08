using MDPActual.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using ExcelDataReader;
using System.Data;
using System.IO;

namespace MDPActual.Controllers
{
    public class HomeController : Controller
    {
        private List<string> names = new List<string>();
        public static string fileName = "";
        private string[] path;
        public static int[] position;

        [HttpGet]
        public IActionResult Index(List<string> items = null)
        {

            names = items == null ? new List<string>() : items;
            return View(names);
        }
        public IActionResult Index(IFormFile file, [FromServices] Microsoft.AspNetCore.Hosting.IHostingEnvironment hostingEnvironment)
        {
            //Для теста
            fileName = $"{hostingEnvironment.WebRootPath}\\files\\{file.FileName}";
            path = Directory.GetCurrentDirectory().Split("\\");
            //Для билда
            /*for (int i = 0; i < path.Length - 3; i++)
                fileName += $"{path[i]}\\";
            fileName += $"{@"\wwwroot\files"}" + "\\" + file.FileName;
            Console.WriteLine(fileName);*/

            using (FileStream fileStream = System.IO.File.Create(fileName))
            {
                file.CopyTo(fileStream);
                fileStream.Flush();
            }
            GetItemList();
            return Index(names);
        }
        private List<string> GetItemList()
        {
            ExcelReader er = new ExcelReader(3, fileName);
            er.SearchNameOfItems(names);
            return names;
            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            //using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            //{
            //    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            //    {
            //        DataTable specificWorkSheet = reader.AsDataSet().Tables[3];
            //        SearchRow sr = new SearchRow();
            //        position = sr.Search(specificWorkSheet, "наименование");
            //        int startRow = position[0];
            //        int startColumn = position[1];
            //        int currentRow = 0;

            //        foreach (var row in specificWorkSheet.Rows)
            //        {
            //            if (currentRow > startRow && ((DataRow)row)[startColumn].ToString() != "")
            //                names.Add(((DataRow)row)[startColumn].ToString());
            //            currentRow++;
            //        }
            //    }
            //}
            //return names;
        }
    }
}