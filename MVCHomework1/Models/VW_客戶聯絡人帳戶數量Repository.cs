using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;

namespace MVCHomework1.Models
{   
	public  class VW_客戶聯絡人帳戶數量Repository : EFRepository<VW_客戶聯絡人帳戶數量>, IVW_客戶聯絡人帳戶數量Repository
	{
        public override IQueryable<VW_客戶聯絡人帳戶數量> All()
        {
            return base.All();
        }
    }

	public  interface IVW_客戶聯絡人帳戶數量Repository : IRepository<VW_客戶聯絡人帳戶數量>
	{

	}
}