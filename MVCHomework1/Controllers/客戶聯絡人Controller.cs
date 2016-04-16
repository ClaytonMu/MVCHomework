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
    public class 客戶聯絡人Controller : BaseController
    {
        //private 客戶資料Entities db = new 客戶資料Entities();
        private 客戶聯絡人Repository _客戶聯絡人Repository = RepositoryHelper.Get客戶聯絡人Repository();
        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();

        public ActionResult Index(SortViewModel viewModel)
        {
            //客戶職稱
            List<string> 職稱List = _客戶聯絡人Repository.All().Select(p => p.職稱).Distinct().OrderBy(p => p).ToList();
            ViewBag.jobTitle = new SelectList(職稱List, viewModel.jobTitle);
            ViewBag.Keyword = viewModel.keyword;
            ViewBag.sortOrder = (viewModel.sortOrder == "ASC") ? "DESC" : "ASC";

            var 客戶聯絡人 = _客戶聯絡人Repository.Sort(viewModel);
            var data = 客戶聯絡人.ToPagedList(viewModel.page, 5);
            return View(data);
        }

        public ActionResult Export(SortViewModel viewModel)
        {
            var 客戶聯絡人 = _客戶聯絡人Repository.Sort(viewModel);
            return File(_客戶聯絡人Repository.Export(客戶聯絡人.ToList()), "application/vnd.ms-excel", "Download.xls");
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
