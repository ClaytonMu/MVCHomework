using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;

namespace MVCHomework1.Models
{   
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{
        

        public 客戶銀行資訊 Find(int id)
        {
            return this.All().Where(p => p.客戶Id == id).FirstOrDefault();
        }

        public override IQueryable<客戶銀行資訊> All()
        {
            var 客戶銀行資訊 = base.All().Include(客 => 客.客戶資料).AsQueryable();
            客戶銀行資訊 = 客戶銀行資訊.Where(客 => 客.是否已刪除 == false);
            return 客戶銀行資訊;
        }

        public IQueryable<客戶銀行資訊> SearchAll(string keyword)
        {
            return base.All().Where(客 =>
                (客.銀行名稱.Contains(keyword) ||
                客.銀行代碼.ToString().Contains(keyword) ||
                客.分行代碼.ToString().Contains(keyword) ||
                客.帳戶名稱.Contains(keyword) ||
                客.客戶資料.客戶名稱.Contains(keyword)) &&
                客.是否已刪除 == false);
        }

        public override void Delete(客戶銀行資訊 entity)
        {
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶銀行資訊Repository : IRepository<客戶銀行資訊>
	{

	}
}