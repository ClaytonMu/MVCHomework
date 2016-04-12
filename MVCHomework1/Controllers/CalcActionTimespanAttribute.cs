using System;
using System.Web.Mvc;

namespace MVCHomework1.Controllers
{
    public class CalcActionTimespanAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.dtStart = DateTime.Now;
            base.OnActionExecuting(filterContext);
        }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {
            var sp = DateTime.Now - (DateTime)filterContext.Controller.ViewBag.dtStart;
            filterContext.Controller.ViewBag.timespan = sp;
            base.OnActionExecuted(filterContext);
        }
    }
}