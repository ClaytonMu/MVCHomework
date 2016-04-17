using System;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using MVCHomework1.Models;
using System.Web.Security;
using System.Collections;
using System.Collections.Generic;

namespace MVCHomework1.Controllers
{
    [Authorize]
    public class AccountController : Controller
    {
        private 客戶資料Repository _客戶repo = RepositoryHelper.Get客戶資料Repository();
        /// <summary>
        /// 設定要存在 FormsAuthenticationTicket 中的資料，這裡用來儲存角色資訊
        /// </summary>
        string userData = "";

        [AllowAnonymous]
        public ActionResult Login()
        {
            return View();
        }

        [AllowAnonymous]
        [HttpPost]
        public ActionResult Login(LoginViewModel loginData)
        {
            if (ModelState.IsValid)
            {
                Session.RemoveAll();

                if (ValidateLogin(loginData))
                {
                    CreateTicket(loginData);
                    return RedirectToAction("Index", "Home");
                }
                else
                {
                    ViewBag.Error = "請確認帳號密碼正確性";
                }
            }
            else
            {
                if (ValidateFirstTimeLogin(loginData))
                {
                    CreateTicket(loginData);
                    return RedirectToAction("Index", "Home");
                }
            }


            return View(loginData);
        }


        private bool ValidateFirstTimeLogin(LoginViewModel loginData)
        {
            var loginUser = _客戶repo.All().Where(p => p.帳號 == loginData.Account).FirstOrDefault();

            if (loginUser != null && loginUser.密碼 == null)
            {
                userData = "customer";
                return true;
            }

            return false;
        }

        private void CreateTicket(LoginViewModel loginData)
        {
            FormsAuthenticationTicket ticket = new FormsAuthenticationTicket(
                        1,
                        loginData.Account,
                        DateTime.Now,
                        DateTime.Now.AddMinutes(30),
                        false,
                        userData,
                        FormsAuthentication.FormsCookiePath);

            string encTicket = FormsAuthentication.Encrypt(ticket);
            Response.Cookies.Add(new HttpCookie(FormsAuthentication.FormsCookieName, encTicket));
        }

        /// <summary>
        /// 驗證使用者是否登入成功
        /// </summary>
        /// <param name="strUsername">登入帳號</param>
        /// <param name="strPassword">登入密碼</param>
        /// <returns></returns>
        private bool ValidateLogin(LoginViewModel loginData)
        {
            // 請自行寫 Code 檢查 Username, Password 是否正確
            SortedList<string, string> sl;
            if (loginData.Account == "admin" && loginData.Password == "123456")
            {
                userData = "admin";
                return true;
            }
            else
            {
                var loginUser = _客戶repo.All().Where(p => p.帳號 == loginData.Account).FirstOrDefault();

                if (loginUser != null)
                {
                    string strPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(loginData.Account + loginData.Password, "MD5");
                    string dbPassword = loginUser.密碼;

                    if (strPassword == dbPassword)
                    {
                        userData = "customer";
                        return true;
                    }
                    else if (dbPassword == null)
                    {
                        userData = "customer";
                        return true;
                    }
                }
            }

            return false;
        }

        [AllowAnonymous]
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login");
        }

        [Authorize(Roles = "customer")]
        public ActionResult EditProfile()
        {
            if (Request.IsAuthenticated)
            {
                var 客戶資料 = _客戶repo.All().Where(p => p.帳號 == User.Identity.Name).FirstOrDefault();
                EditProfileViewModel model = new EditProfileViewModel();
                model.電話 = 客戶資料.電話;
                model.傳真 = 客戶資料.傳真;
                model.地址 = 客戶資料.地址;
                model.Email = 客戶資料.Email;
                model.密碼 = "";
                客戶資料.密碼 = "";
                return View(model);

            }
            else
            {
                return RedirectToAction("Login");
            }
            
        }

        [Authorize(Roles = "customer")]
        [HttpPost]
        public ActionResult EditProfile(EditProfileViewModel model)
        {
            var 客戶資料 = _客戶repo.All().Where(p => p.帳號 == User.Identity.Name).FirstOrDefault();
            if (ModelState.IsValid)
            {
                客戶資料.電話 = model.電話;
                客戶資料.傳真 = model.傳真;
                客戶資料.地址 = model.地址;
                客戶資料.Email = model.Email;

                if(!string.IsNullOrEmpty(model.確認密碼))
                    客戶資料.密碼 = FormsAuthentication.HashPasswordForStoringInConfigFile(客戶資料.帳號 + model.確認密碼, "MD5");

                _客戶repo.UnitOfWork.Commit();
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(model);
            }

            
        }
    }
}