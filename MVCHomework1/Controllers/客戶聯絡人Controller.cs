using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVCHomework1.Models;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace MVCHomework1.Controllers
{
    public class 客戶聯絡人Controller : Controller
    {
        //private 客戶資料Entities db = new 客戶資料Entities();
        private 客戶聯絡人Repository _客戶聯絡人Repository = RepositoryHelper.Get客戶聯絡人Repository();
        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶聯絡人
        public ActionResult Index()
        {
            var 客戶聯絡人 = _客戶聯絡人Repository.All();

            return View(客戶聯絡人.ToList());
        }

        [HttpPost]
        public ActionResult Index(string keyword)
        {
            var 客戶聯絡人 = _客戶聯絡人Repository.SearchAll(keyword);

            return View(客戶聯絡人.ToList());
        }

        public ActionResult ExportToExcel()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("職稱");
            dt.Columns.Add("姓名");
            dt.Columns.Add("Email");
            dt.Columns.Add("手機");
            dt.Columns.Add("電話");
            dt.Columns.Add("客戶名稱");

            _客戶聯絡人Repository.All().ToList().ForEach(p => {
                
                DataRow dr = dt.NewRow();
                dr["職稱"] = p.職稱;
                dr["姓名"] = p.姓名;
                dr["Email"] = p.Email;
                dr["手機"] = p.手機;
                dr["電話"] = p.電話;
                dr["客戶名稱"] = p.客戶資料.客戶名稱;
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
            return File(MS.ToArray(), "application/vnd.ms-excel", "Download.xls");
        }

        // GET: 客戶聯絡人/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = _客戶聯絡人Repository.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶聯絡人/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                _客戶聯絡人Repository.Add(客戶聯絡人);
                _客戶聯絡人Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = _客戶聯絡人Repository.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                _客戶聯絡人Repository.UnitOfWork.Context.Entry(客戶聯絡人).State = EntityState.Modified;
                _客戶聯絡人Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        

        // GET: 客戶聯絡人/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = _客戶聯絡人Repository.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            var 聯絡人 = _客戶聯絡人Repository.Find(id);
            _客戶聯絡人Repository.Delete(聯絡人);

            _客戶聯絡人Repository.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _客戶聯絡人Repository.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
