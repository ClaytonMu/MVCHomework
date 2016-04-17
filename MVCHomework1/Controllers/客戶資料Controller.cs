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
using PagedList;

namespace MVCHomework1.Controllers
{
    public class 客戶資料Controller : BaseController
    {

        private 客戶資料Repository _客戶資料Repository = RepositoryHelper.Get客戶資料Repository();
        private 客戶聯絡人Repository _客戶聯絡人Repository = RepositoryHelper.Get客戶聯絡人Repository();
        private VW_客戶聯絡人帳戶數量Repository _VW_客戶聯絡人帳戶數量Repository = RepositoryHelper.GetVW_客戶聯絡人帳戶數量Repository();

        public ActionResult TestError()
        {
            throw new ArgumentException("錯誤");
        }

        // GET: 客戶資料
        public ActionResult Index(SortViewModel viewModel)
        {
            ViewBag.Keyword = viewModel.keyword;
            ViewBag.sortOrder = (viewModel.sortOrder == "ASC") ? "DESC" : "ASC";

            var 客戶資料 = _客戶資料Repository.Sort(viewModel);
            var data = 客戶資料.ToPagedList(viewModel.page, 5);
            return View(data);
        }

        public ActionResult Export(SortViewModel viewModel)
        {
            var 客戶資料 = _客戶資料Repository.Sort(viewModel);
            return File(_客戶資料Repository.Export(客戶資料.ToList()), "application/vnd.ms-excel", "Download.xls");
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
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶類別")] 客戶資料 客戶資料)
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

            ViewBag.客戶類別 = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value", 客戶資料.客戶類別);
            return View(客戶資料);
        }

        [HttpPost]
        public ActionResult BatchEdit(IList<客戶聯絡人> 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                foreach (var item in 客戶聯絡人)
                {
                    var 聯絡人 = _客戶聯絡人Repository.Find(item.Id);
                    聯絡人.職稱 = item.職稱;
                    聯絡人.手機 = item.手機;
                    聯絡人.電話 = item.電話;
                }

                _客戶聯絡人Repository.UnitOfWork.Commit();

                return RedirectToAction("Index");
            }
            else
            {
                return View("Detail");
            }
            
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(客戶資料 客資)
        {
            var 客戶資料 = _客戶資料Repository.Find(客資.Id);

            if (TryUpdateModel(客戶資料, new string[] { "Id", "客戶名稱", "統一編號", "電話", "傳真", "地址", "Email", "客戶類別", "帳號" }))
            {
                _客戶資料Repository.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }


            ViewBag.客戶類別 = new SelectList(_客戶資料Repository.所有客戶類別(), "Key", "Value", 客資.客戶類別);
            return View(客資);
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

        public ActionResult MVCGrid()
        {
            return View(_VW_客戶聯絡人帳戶數量Repository.All().ToList());
        }

        public ActionResult 客戶聯絡人銀行帳戶數量()
        {
            return View(_VW_客戶聯絡人帳戶數量Repository.All().ToList());
        }

        public ActionResult JsonTest(int? id)
        {
            _客戶資料Repository.UnitOfWork.Context.Configuration.LazyLoadingEnabled = false;
            var data = _客戶資料Repository.Find(id.Value);
            data.密碼 = "";
            
            return Json(data, JsonRequestBehavior.AllowGet);
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


