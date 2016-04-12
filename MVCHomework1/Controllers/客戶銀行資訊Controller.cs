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
    public class 客戶銀行資訊Controller : BaseController
    {
        //private 客戶資料Entities db = new 客戶資料Entities();
        private 客戶銀行資訊Repository _客戶銀行資訊Repository = RepositoryHelper.Get客戶銀行資訊Repository();
        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶銀行資訊
        public ActionResult Index(string sortOrder = "", string keyword = "")
        {
            var 客戶銀行資訊 = _客戶銀行資訊Repository.All().OrderBy(p => p.帳戶名稱).AsQueryable();

            if (!string.IsNullOrEmpty(keyword))
            {
                客戶銀行資訊 = _客戶銀行資訊Repository.SearchAll(keyword);
            }


            if (!string.IsNullOrEmpty(sortOrder))
            {
                string name = sortOrder.Split('_')[0];
                string order = sortOrder.Split('_')[1];

                if (order == "ASC")
                {
                    order = "DESC";
                    ViewBag.SortOrder = "DESC";
                }
                else
                {
                    order = "ASC";
                    ViewBag.SortOrder = "ASC";
                }

                switch (name)
                {
                    case "銀行名稱":
                        客戶銀行資訊 = order == "ASC" ? 客戶銀行資訊.OrderBy(p => p.銀行名稱) : 客戶銀行資訊.OrderByDescending(p => p.銀行名稱);
                        break;
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
                    default:
                        break;
                }
            }
            else
            {
                //客戶聯絡人 = _客戶聯絡人Repository.All().OrderBy(p => p.姓名);
                ViewBag.SortOrder = "ASC";
                ViewBag.Keyword = "";
                //return View(客戶聯絡人.ToList());
            }

            return View(客戶銀行資訊.ToList());


            
        }

        [HttpPost]
        public ActionResult Index(string keyword)
        {
            var 客戶銀行資訊 = _客戶銀行資訊Repository.SearchAll(keyword);

            return View(客戶銀行資訊.ToList());
        }

        public ActionResult ExportToExcel()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("銀行名稱");
            dt.Columns.Add("銀行代碼");
            dt.Columns.Add("分行代碼");
            dt.Columns.Add("帳戶名稱");
            dt.Columns.Add("帳戶號碼");
            dt.Columns.Add("客戶名稱");

            _客戶銀行資訊Repository.All().ToList().ForEach(p => {
                DataRow dr = dt.NewRow();
                dr["銀行名稱"] = p.銀行名稱;
                dr["銀行代碼"] = p.銀行代碼;
                dr["分行代碼"] = p.分行代碼;
                dr["帳戶名稱"] = p.帳戶名稱;
                dr["帳戶號碼"] = p.帳戶號碼;
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

        // GET: 客戶銀行資訊/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = _客戶銀行資訊Repository.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶銀行資訊/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                _客戶銀行資訊Repository.Add(客戶銀行資訊);
                _客戶銀行資訊Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = _客戶銀行資訊Repository.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(_客戶資料Repository.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                _客戶銀行資訊Repository.UnitOfWork.Context.Entry(客戶銀行資訊).State = EntityState.Modified;
                _客戶銀行資訊Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(((客戶資料Entities)_客戶銀行資訊Repository.UnitOfWork.Context).客戶資料, "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = _客戶銀行資訊Repository.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            var 銀行資訊 = _客戶銀行資訊Repository.Find(id);
            _客戶銀行資訊Repository.Delete(銀行資訊);

            _客戶銀行資訊Repository.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _客戶銀行資訊Repository.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
