﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCHomework1.Controllers
{
    [HandleError(ExceptionType = typeof(ArgumentException), View = "CustomError")]
    [CalcActionTimespan]
    [Authorize(Roles = "admin")]
    public class BaseController : Controller
    {

    }
}