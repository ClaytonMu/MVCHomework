using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MVCHomework1.Models
{
    public class EditProfileViewModel
    {
        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [Required]
        public string 電話 { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        public string 傳真 { get; set; }

        [StringLength(100, ErrorMessage = "欄位長度不得大於 100 個字元")]
        public string 地址 { get; set; }

        [StringLength(250, ErrorMessage = "欄位長度不得大於 250 個字元")]
        [EmailAddress(ErrorMessage = "無效的Email格式")]
        public string Email { get; set; }

        [DataType(DataType.Password)]
        public string 密碼 { get; set; }

        [DataType(DataType.Password)]
        [Compare("密碼", ErrorMessage = "密碼和確認密碼不相符。")]
        public string 確認密碼 { get; set; }
    }
}