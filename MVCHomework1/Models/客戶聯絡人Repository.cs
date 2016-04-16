using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;

namespace MVCHomework1.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        private Utility _utility = new Utility();
        public 客戶聯絡人 Find(int id)
        {
            return this.All().Where(p => p.Id == id).FirstOrDefault();
        }

        public override IQueryable<客戶聯絡人> All()
        {
            var 客戶聯絡人 = base.All().Include(客 => 客.客戶資料).AsQueryable();
            客戶聯絡人 = 客戶聯絡人.Where(客 => 客.是否已刪除 == false);

            return 客戶聯絡人;
        }

        public IQueryable<客戶聯絡人> SearchKeyword(string keyword)
        {
            return base.All().Where(客 =>
                (客.姓名.Contains(keyword) ||
                客.職稱.Contains(keyword) ||
                客.Email.Contains(keyword) ||
                客.手機.Contains(keyword) ||
                客.客戶資料.客戶名稱.Contains(keyword)) &&
                客.是否已刪除 == false);
        }

        public IQueryable<客戶聯絡人> Sort(SortViewModel sortModel)
        {
            IQueryable<客戶聯絡人> 客戶聯絡人 = null;

            if (!string.IsNullOrEmpty(sortModel.keyword))
            {
                客戶聯絡人 = this.SearchKeyword(sortModel.keyword);
            }
            else
            {
                客戶聯絡人 = this.All();
            }

            if (!string.IsNullOrEmpty(sortModel.jobTitle))
            {
                客戶聯絡人 = 客戶聯絡人.Where(p => p.職稱 == sortModel.jobTitle);
            }

            if (!string.IsNullOrEmpty(sortModel.sortOrder))
            {
                string order = (sortModel.sortOrder == "ASC") ? "DESC" : "ASC";

                switch (sortModel.sortColumn)
                {
                    case "職稱":
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.職稱) : 客戶聯絡人.OrderByDescending(p => p.職稱);
                        break;
                    
                    case "Email":
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.Email) : 客戶聯絡人.OrderByDescending(p => p.Email);
                        break;
                    case "手機":
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.手機) : 客戶聯絡人.OrderByDescending(p => p.手機);
                        break;
                    case "電話":
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.電話) : 客戶聯絡人.OrderByDescending(p => p.電話);
                        break;
                    case "客戶名稱":
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.客戶資料.客戶名稱) : 客戶聯絡人.OrderByDescending(p => p.客戶資料.客戶名稱);
                        break;
                    case "姓名":
                    default:
                        客戶聯絡人 = order == "ASC" ? 客戶聯絡人.OrderBy(p => p.姓名) : 客戶聯絡人.OrderByDescending(p => p.姓名);
                        break;
                }
            }

            return 客戶聯絡人;
        }

        public byte[] Export(IList<客戶聯絡人> 客戶聯絡人)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("職稱");
            dt.Columns.Add("姓名");
            dt.Columns.Add("Email");
            dt.Columns.Add("手機");
            dt.Columns.Add("電話");
            dt.Columns.Add("客戶名稱");

            客戶聯絡人.ToList().ForEach(p => {

                DataRow dr = dt.NewRow();
                dr["職稱"] = p.職稱;
                dr["姓名"] = p.姓名;
                dr["Email"] = p.Email;
                dr["手機"] = p.手機;
                dr["電話"] = p.電話;
                dr["客戶名稱"] = p.客戶資料.客戶名稱;
                dt.Rows.Add(dr);
            });

            return _utility.Export(dt);
        }

        public override void Delete(客戶聯絡人 entity)
        {
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}