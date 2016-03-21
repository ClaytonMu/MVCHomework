namespace MVCHomework1.Models
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;

    [MetadataType(typeof(客戶聯絡人MetaData))]
    public partial class 客戶聯絡人 : IValidatableObject
    {
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            客戶資料Entities db = new 客戶資料Entities();

            if (this.Id == default(int))
            {
                //新增
                var count = db.客戶聯絡人.Count(客 =>
                    客.客戶Id == this.客戶Id &&
                    客.Email == this.Email &&
                    客.是否已刪除 == false
                );

                if (count > 0)
                {
                    yield return new ValidationResult("Email已存在", new string[] { "Email" });
                }
            }
            else
            {
                var count = db.客戶聯絡人.Count(客 =>
                    客.客戶Id == this.客戶Id &&
                    客.Email == this.Email &&
                    客.Id != this.Id &&
                    客.是否已刪除 == false
                );

                if (count > 0)
                {
                    yield return new ValidationResult("Email已存在", new string[] { "Email" });
                }

            }
        }
    }

    public partial class 客戶聯絡人MetaData
    {
        [Required]
        public int Id { get; set; }
        [Required]
        public int 客戶Id { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        [Required]
        public string 職稱 { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        [Required]
        public string 姓名 { get; set; }
        
        [StringLength(250, ErrorMessage="欄位長度不得大於 250 個字元")]
        [Required]
        [EmailAddress(ErrorMessage = "無效的Email格式")]
        public string Email { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        [驗證手機(ErrorMessage = "無效的手機格式")]
        public string 手機 { get; set; }
        
        [StringLength(50, ErrorMessage="欄位長度不得大於 50 個字元")]
        public string 電話 { get; set; }
        [Required]
        public bool 是否已刪除 { get; set; }
    
        public virtual 客戶資料 客戶資料 { get; set; }
    }
}
