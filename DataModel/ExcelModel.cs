using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataModel
{
    public class ExcelModel
    {
        public int RId { get; set; }
        public int CID { get; set; }              //客户ID
        public string CName { get; set; }            //客户名称
        public string BNO { get; set; }           //订单号
        public DateTime BTime { get; set; }       //下单时间
        public string Product { get; set; }       //商品名称
        public int PNum { get; set; }             //数量
        public decimal TotalMoney { get; set; }   //总金额
        public decimal Discount { get; set; }     //折扣
        public string Remark { get; set; }        //备注
    }
}