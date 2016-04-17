using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace MVCHomework1.Models
{   
	public  class 客戶資料Repository : EFRepository<客戶資料>, I客戶資料Repository
	{
        private Utility _utility = new Utility();

        public enum 客戶類別
        {
            無,
            營造,
            食品相關,
            電子
        }

        public List<客戶類別ViewModel> 所有客戶類別()
        {
            List<客戶類別ViewModel> list = new List<客戶類別ViewModel>(); 
            foreach (客戶類別 type in Enum.GetValues(typeof(客戶類別)))
            {
                客戶類別ViewModel m = new 客戶類別ViewModel();
                m.Key = type.ToString();
                m.Value = type.ToString();
                list.Add(m);
            }

            return list;
        }

        public 客戶資料 Find(int id)
        {
            return this.All().Where(p => p.Id == id && p.是否已刪除 == false).FirstOrDefault();
        }

        public override IQueryable<客戶資料> All()
        {
            //var a = base.All().Include(p => p.客戶聯絡人).Include(p => p.客戶銀行資訊);

            //base.All().Where(客 => 客.是否已刪除 == false).ToList().ForEach(
            //        客 => { 客.客戶類別 = ((客戶類別)Enum.Parse(typeof(客戶類別), 客.客戶類別)).ToString(); }
            //    );



            return base.All().Where(客 => 客.是否已刪除 == false);
        }

        public IQueryable<客戶資料> SearchKeyword(string keyword)
        {
            return this.All().Where(客 =>
                客.客戶名稱.Contains(keyword) ||
                客.統一編號.Contains(keyword) ||
                客.電話.Contains(keyword) ||
                客.傳真.Contains(keyword) ||
                客.地址.Contains(keyword) ||
                客.Email.Contains(keyword) ||
                客.客戶類別.Contains(keyword));
        }

        public IQueryable<客戶資料> Sort(SortViewModel sortModel)
        {
            IQueryable<客戶資料> 客戶資料 = null;
            if (!string.IsNullOrEmpty(sortModel.keyword))
            {
                客戶資料 = this.SearchKeyword(sortModel.keyword);
            }
            else
            {
                客戶資料 = this.All();
            }

            if (!string.IsNullOrEmpty(sortModel.sortOrder))
            {
                string order = sortModel.sortOrder;
                switch (sortModel.sortColumn)
                {
                    
                    case "統一編號":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.統一編號) : 客戶資料.OrderByDescending(p => p.統一編號);
                        break;
                    case "電話":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.電話) : 客戶資料.OrderByDescending(p => p.電話);
                        break;
                    case "傳真":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.傳真) : 客戶資料.OrderByDescending(p => p.傳真);
                        break;
                    case "地址":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.地址) : 客戶資料.OrderByDescending(p => p.地址);
                        break;
                    case "Email":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.Email) : 客戶資料.OrderByDescending(p => p.Email);
                        break;
                    case "客戶類別":
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.客戶類別) : 客戶資料.OrderByDescending(p => p.客戶類別);
                        break;
                    case "客戶名稱":
                    default:
                        客戶資料 = order == "ASC" ? 客戶資料.OrderBy(p => p.客戶名稱) : 客戶資料.OrderByDescending(p => p.客戶名稱);
                        break;
                }
            }

            return 客戶資料;
        }

        public override void Delete(客戶資料 entity)
        {
            entity.是否已刪除 = true;
            entity.客戶聯絡人.Where(客 => 客.客戶Id == entity.Id)?.ToList().ForEach(客 => { 客.是否已刪除 = true; });
            entity.客戶銀行資訊.Where(客 => 客.客戶Id == entity.Id)?.ToList().ForEach(客 => { 客.是否已刪除 = true; });
        }

        public byte[] Export(IList<客戶資料> 客戶資料)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("客戶名稱");
            dt.Columns.Add("統一編號");
            dt.Columns.Add("電話");
            dt.Columns.Add("傳真");
            dt.Columns.Add("地址");
            dt.Columns.Add("Email");
            dt.Columns.Add("客戶類別");

            客戶資料.ToList().ForEach(p => {

                DataRow dr = dt.NewRow();
                dr["客戶名稱"] = p.客戶名稱;
                dr["統一編號"] = p.統一編號;
                dr["電話"] = p.電話;
                dr["傳真"] = p.傳真;
                dr["地址"] = p.地址;
                dr["Email"] = p.Email;
                dr["客戶類別"] = p.客戶類別;
                dt.Rows.Add(dr);
            });

            return _utility.Export(dt);
        }
    }

	public  interface I客戶資料Repository : IRepository<客戶資料>
	{

	}
}