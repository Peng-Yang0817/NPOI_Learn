using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NPOI_Learn.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public void Export_xlsx()
        {
            //創建xlsx物件
            XSSFWorkbook wb = new XSSFWorkbook();
            //在wb內建立一張sheet
            var sheet = wb.CreateSheet("table");
            //創建row索引，先不給予賦值
            IRow row = null;
            //對row的個別cell進行賦值
            for (int i = 0; i < 5; i++)
            {
                row = sheet.CreateRow(i);
                for (int j = 0; j < 4; j++)
                {
                    row.CreateCell(j).SetCellValue("妳好" + "+" + i + "+" + j);
                }
            }
            //creatRow即表示初始化Row
            //sheet.CreateRow(0);

            //利用memoryStream，將wb物件轉為byte[] array
            MemoryStream ms = new MemoryStream();
            wb.Write(ms);
            byte[] bytes = ms.ToArray();
            ms.Close();

            //對Response域賦值
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "aaa.xlsx");
            Response.BinaryWrite(bytes);
            Response.End();
        }
    }
}