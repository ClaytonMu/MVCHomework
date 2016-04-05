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
                m.Key = (int)type;
                m.Value = type.ToString();
                list.Add(m);
            }

            return list;
        }

        public 客戶資料 Find(int id)
        {
            return base.All().Where(p => p.Id == id && p.是否已刪除 == false).FirstOrDefault();
        }

        public override IQueryable<客戶資料> All()
        {
            base.All().Where(客 => 客.是否已刪除 == false).ToList().ForEach(
                    客 => { 客.客戶類別 = ((客戶類別)Enum.Parse(typeof(客戶類別), 客.客戶類別)).ToString(); }
                );

            return base.All().Where(客 => 客.是否已刪除 == false);
        }

        public IQueryable<客戶資料> SearchAll(string keyword)
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

        public override void Delete(客戶資料 entity)
        {
            entity.是否已刪除 = true;
            entity.客戶聯絡人.Where(客 => 客.客戶Id == entity.Id)?.ToList().ForEach(客 => { 客.是否已刪除 = true; });
            entity.客戶銀行資訊.Where(客 => 客.客戶Id == entity.Id)?.ToList().ForEach(客 => { 客.是否已刪除 = true; });
        }

        public byte[] ExportToExcel()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("客戶名稱");
            dt.Columns.Add("統一編號");
            dt.Columns.Add("電話");
            dt.Columns.Add("傳真");
            dt.Columns.Add("地址");
            dt.Columns.Add("Email");
            dt.Columns.Add("客戶類別");

            this.All().ToList().ForEach(p => {

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

            //建立Excel 2003檔案
            IWorkbook wb = new HSSFWorkbook();
            ISheet ws;

            ////建立Excel 2007檔案
            //IWorkbook wb = new XSSFWorkbook();
            //ISheet ws;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            MemoryStream MS = new MemoryStream();
            wb.Write(MS);
            MS.Close();
            MS.Dispose();
            return MS.ToArray();
        }
    }

	public  interface I客戶資料Repository : IRepository<客戶資料>
	{

	}
}