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
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{

        private Utility _utility = new Utility();
        public 客戶銀行資訊 Find(int id)
        {
            return this.All().Where(p => p.Id == id).FirstOrDefault();
        }

        public override IQueryable<客戶銀行資訊> All()
        {
            var 客戶銀行資訊 = base.All().Include(客 => 客.客戶資料).AsQueryable();
            客戶銀行資訊 = 客戶銀行資訊.Where(客 => 客.是否已刪除 == false);
            return 客戶銀行資訊;
        }

        public IQueryable<客戶銀行資訊> SearchKeyword(string keyword)
        {
            return base.All().Where(客 =>
                (客.銀行名稱.Contains(keyword) ||
                客.銀行代碼.ToString().Contains(keyword) ||
                客.分行代碼.ToString().Contains(keyword) ||
                客.帳戶名稱.Contains(keyword) ||
                客.客戶資料.客戶名稱.Contains(keyword)) &&
                客.是否已刪除 == false);
        }

        public IQueryable<客戶銀行資訊> Sort(SortViewModel sortModel)
        {
            IQueryable<客戶銀行資訊> 客戶銀行資訊 = null;
            if (!string.IsNullOrEmpty(sortModel.keyword))
            {
                客戶銀行資訊 = this.SearchKeyword(sortModel.keyword);
            }
            else
            {
                客戶銀行資訊 = this.All();
            }

            if (!string.IsNullOrEmpty(sortModel.sortOrder))
            {
                string order = sortModel.sortOrder;
                switch (sortModel.sortColumn)
                {
                    
                    case "銀行代碼":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.銀行代碼) : 客戶銀行資訊.OrderByDescending(p => p.銀行代碼);
                        break;
                    case "分行代碼":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.分行代碼) : 客戶銀行資訊.OrderByDescending(p => p.分行代碼);
                        break;
                    case "帳戶名稱":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.帳戶名稱) : 客戶銀行資訊.OrderByDescending(p => p.帳戶名稱);
                        break;
                    case "帳戶號碼":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.帳戶號碼) : 客戶銀行資訊.OrderByDescending(p => p.帳戶號碼);
                        break;
                    case "客戶名稱":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.客戶資料.客戶名稱) : 客戶銀行資訊.OrderByDescending(p => p.客戶資料.客戶名稱);
                        break;
                    case "銀行名稱":
                    default:
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.銀行名稱) : 客戶銀行資訊.OrderByDescending(p => p.銀行名稱);
                        break;
                }
            }

            return 客戶銀行資訊;
        }

        public byte[] Export(IList<客戶銀行資訊> 客戶銀行資訊)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("銀行名稱");
            dt.Columns.Add("銀行代碼");
            dt.Columns.Add("分行代碼");
            dt.Columns.Add("帳戶名稱");
            dt.Columns.Add("帳戶號碼");
            dt.Columns.Add("客戶名稱");

            客戶銀行資訊.ToList().ForEach(p => {
                DataRow dr = dt.NewRow();
                dr["銀行名稱"] = p.銀行名稱;
                dr["銀行代碼"] = p.銀行代碼;
                dr["分行代碼"] = p.分行代碼;
                dr["帳戶名稱"] = p.帳戶名稱;
                dr["帳戶號碼"] = p.帳戶號碼;
                dr["客戶名稱"] = p.客戶資料.客戶名稱;
                dt.Rows.Add(dr);
            });

            return _utility.Export(dt);
        }


        public override void Delete(客戶銀行資訊 entity)
        {
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶銀行資訊Repository : IRepository<客戶銀行資訊>
	{

	}
}