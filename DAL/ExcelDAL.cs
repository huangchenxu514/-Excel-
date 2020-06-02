using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataModel;

namespace DAL
{
    public class ExcelDAL
    {
        DBHelper db = new DBHelper();
        public List<ExcelModel> ShowBuy_Product(string Name = "")
        {
            string str = "select * from Buy_Product b join Customer c on b.CID = c.CId where 1 = 1 ";
            if (!string.IsNullOrEmpty(Name))
            {
                str += " and c.CName like '%" + Name + "%'";
            }
            return db.GetToList<ExcelModel>(str);
        }
    }
}
