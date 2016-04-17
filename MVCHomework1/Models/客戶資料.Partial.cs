namespace MVCHomework1.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;

    [MetadataType(typeof(客戶資料MetaData))]
    public partial class 客戶資料 : IValidatableObject
    {
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            客戶資料Entities db = new 客戶資料Entities();

            if (this.Id == default(int))
            {
                //新增
                var count = db.客戶資料.Count(客 =>
                    客.帳號.ToLower() == this.Email.ToLower() &&
                    客.是否已刪除 == false
                );

                if (count > 0)
                {
                    yield return new ValidationResult("帳號已存在", new string[] { "帳號" });
                }
            }
            else
            {
                var count = db.客戶資料.Count(客 =>
                    客.帳號.ToLower() == this.帳號.ToLower() &&
                    客.Id != this.Id &&
                    客.是否已刪除 == false
                );

                if (count > 0)
                {
                    yield return new ValidationResult("帳號已存在", new string[] { "帳號" });
                }

            }
        }
    }

    public partial class 客戶資料MetaData
    {
        [Required]
        public int Id { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        [Required]
        public string 客戶名稱 { get; set; }
        
        [StringLength(8, ErrorMessage="欄位長度不得大於 8 個字元")]
        [Required]
        public string 統一編號 { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        [Required]
        public string 電話 { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        public string 傳真 { get; set; }
        
        [StringLength(100, ErrorMessage="欄位長度不得大於 100 個字元")]
        public string 地址 { get; set; }
        
        [StringLength(250, ErrorMessage="欄位長度不得大於 250 個字元")]
        [EmailAddress(ErrorMessage = "無效的Email格式")]
        public string Email { get; set; }
        [Required]
        public bool 是否已刪除 { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        public string 客戶類別 { get; set; }

        public string 帳號 { get; set; }

        [DataType(DataType.Password)]
        public string 密碼 { get; set; }

        public virtual ICollection<客戶銀行資訊> 客戶銀行資訊 { get; set; }
        public virtual ICollection<客戶聯絡人> 客戶聯絡人 { get; set; }
    }
}
