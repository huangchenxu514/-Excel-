using DAL;
using DataModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication1.Controllers
{
    public class ExcelController : Controller
    {
        ExcelDAL dal = new ExcelDAL();
        // GET: Excel
        public ActionResult Index()
        {
            // return View(dal.ShowBuy_Product());
            return View();
        }
        public ActionResult ExportExcel(/*FormCollection form*/)
        {
            //创建数据
            List<ExcelModel> models = dal.ShowBuy_Product();

            //定义返回的stream流
            MemoryStream stream = new MemoryStream();
            //创建一个workbook
            var workbook = new HSSFWorkbook();
            //创建一个sheet
            var sheet = workbook.CreateSheet("工作表格");
            //创建一个空行，作为标题
            var headRow = sheet.CreateRow(0);
            headRow.CreateCell(0).SetCellValue("客户姓名");
            headRow.CreateCell(1).SetCellValue("订单号");
            headRow.CreateCell(2).SetCellValue("下单时间");
            headRow.CreateCell(3).SetCellValue("商品名称");
            headRow.CreateCell(4).SetCellValue("数量");
            headRow.CreateCell(5).SetCellValue("总金额");
            headRow.CreateCell(6).SetCellValue("折扣");
            headRow.CreateCell(7).SetCellValue("备注");

            //创建内容行
            var index = 1;
            foreach (var item in models)
            {
                var newRow = sheet.CreateRow(index);
                newRow.CreateCell(0).SetCellValue(item.CName);
                newRow.CreateCell(1).SetCellValue(item.BNO);
                newRow.CreateCell(2).SetCellValue(item.BTime);
                newRow.CreateCell(3).SetCellValue(item.Product);
                newRow.CreateCell(4).SetCellValue(item.PNum);
                newRow.CreateCell(5).SetCellValue(Convert.ToDouble(item.TotalMoney));
                newRow.CreateCell(6).SetCellValue(Convert.ToDouble(item.Discount));
                newRow.CreateCell(7).SetCellValue(item.Remark);
                index++;
            }
            //将Excel内容写入stream流
            workbook.Write(stream);
            //清理资源
            stream.Flush();
            stream.Position = 0;
            //sheet = null;
            //headRow = null;
            //workbook = null;
            //返回
            return File(stream, "application/vnd.ms-excel", "导出的Excel文档.xls");
        }
    }
}