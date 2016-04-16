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
using PagedList;

namespace MVCHomework1.Controllers
{
    public class 客戶銀行資訊Controller : BaseController
    {
        //private 客戶資料Entities db = new 客戶資料Entities();
        private 客戶銀行資訊Repository _客戶銀行資訊Repository = RepositoryHelper.Get客戶銀行資訊Repository();
        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶銀行資訊
        public ActionResult Index(SortViewModel viewModel)
        {
            ViewBag.Keyword = viewModel.keyword;
            ViewBag.sortOrder = (viewModel.sortOrder == "ASC") ? "DESC" : "ASC";

            var 客戶銀行資訊 = _客戶銀行資訊Repository.Sort(viewModel);
            var data = 客戶銀行資訊.ToPagedList(viewModel.page, 5);
            return View(data);
        }

        public ActionResult Export(SortViewModel viewModel)
        {
            var 客戶銀行資訊 = _客戶銀行資訊Repository.Sort(viewModel);
            return File(_客戶銀行資訊Repository.Export(客戶銀行資訊.ToList()), "application/vnd.ms-excel", "Download.xls");
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
