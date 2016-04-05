using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;

namespace MVCHomework1.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        private 客戶資料Entities db = (客戶資料Entities)RepositoryHelper.Get客戶資料Repository().UnitOfWork.Context;

        public 客戶聯絡人 Find(int id)
        {
            return this.All().Where(p => p.Id == id).FirstOrDefault();
        }

        public override IQueryable<客戶聯絡人> All()
        {
            var 客戶聯絡人 = base.All().Include(客 => 客.客戶資料).AsQueryable();
            客戶聯絡人 = 客戶聯絡人.Where(客 => 客.是否已刪除 == false);

            return 客戶聯絡人;
        }

        public IQueryable<客戶聯絡人> SearchAll(string keyword)
        {
            return base.All().Where(客 =>
                (客.姓名.Contains(keyword) ||
                客.職稱.Contains(keyword) ||
                客.Email.Contains(keyword) ||
                客.手機.Contains(keyword) ||
                客.客戶資料.客戶名稱.Contains(keyword)) &&
                客.是否已刪除 == false);
        }

        public override void Delete(客戶聯絡人 entity)
        {
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}