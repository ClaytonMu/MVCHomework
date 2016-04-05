using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MVCHomework1.Models
{
    public class 客戶聯絡人ViewModel
    {
        [Required]
        public int Id { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [Required]
        public string 職稱 { get; set; }


        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [驗證手機(ErrorMessage = "無效的手機格式")]
        public string 手機 { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        public string 電話 { get; set; }

    }
}