using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVCHomework1.Models;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;

namespace MVCHomework1.Controllers
{
    public class 客戶資料Controller : Controller
    {

        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();
        private 客戶聯絡人Repository _客戶聯絡人Repository = RepositoryHelper.Get客戶聯絡人Repository();
        private VW_客戶聯絡人帳戶數量Repository _VW_客戶聯絡人帳戶數量Repository = RepositoryHelper.GetVW_客戶聯絡人帳戶數量Repository();
        // GET: 客戶資料
        public ActionResult Index()
        {
            var 客戶資料 = _客戶資料Repository.All();

            return View(客戶資料.ToList());
        }

        [HttpPost]
        public ActionResult Index(string keyword)
        {
            var 客戶資料 = _客戶資料Repository.SearchAll(keyword);

            return View(客戶資料.ToList());
        }

        public ActionResult ExportToExcel()
        {
            byte[] info = _客戶資料Repository.ExportToExcel();
            return File(info, "application/vnd.ms-excel", "Download.xls");
        }

        // GET: 客戶資料/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = _客戶資料Repository.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Create
        public ActionResult Create()
        {
            ViewBag.類別Id = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value");
            return View();
        }

        // POST: 客戶資料/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                _客戶資料Repository.Add(客戶資料);
                _客戶資料Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.類別Id = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value");
            return View(客戶資料);
        }

        public ActionResult GetContracterPartial(IList<客戶聯絡人> 客戶聯絡人)
        {
            //客戶聯絡人Repository 客戶聯絡人Repository = RepositoryHelper.Get客戶聯絡人Repository();
            return PartialView("_IndexPartial", 客戶聯絡人);
        }

        [HttpPost]
        public ActionResult EditPartial(IList<客戶聯絡人ViewModel> data)
        {
            if (ModelState.IsValid)
            {
                foreach (var item in data)
                {
                    var 客戶聯絡人 = _客戶聯絡人Repository.Find(item.Id);
                    客戶聯絡人.職稱 = item.職稱;
                    客戶聯絡人.電話 = item.電話;
                    客戶聯絡人.手機 = item.手機;
                }

                _客戶聯絡人Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            else
            {
                return RedirectToAction("Index");
            }
            //return PartialView("_IndexPartial", _客戶資料Repository.Find(id));
        }

        // GET: 客戶資料/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = _客戶資料Repository.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }

            ViewBag.類別Id = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value");
            return View(客戶資料);
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶類別,客戶聯絡人")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                _客戶資料Repository.UnitOfWork.Context.Entry(客戶資料).State = EntityState.Modified;
                _客戶資料Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.類別Id = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value");
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = _客戶資料Repository.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }

            return View(客戶資料);
        }

        // POST: 客戶資料/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = _客戶資料Repository.Find(id);

            _客戶資料Repository.Delete(客戶資料);

            _客戶資料Repository.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        public ActionResult 客戶聯絡人銀行帳戶數量()
        {
            return View(_VW_客戶聯絡人帳戶數量Repository.All().ToList());
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _客戶資料Repository.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}


