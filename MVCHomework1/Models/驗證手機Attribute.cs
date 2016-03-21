using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace MVCHomework1.Models
{
    public class 驗證手機Attribute : DataTypeAttribute
    {
        public 驗證手機Attribute() : base(DataType.PhoneNumber)
        {

        }
        public override bool IsValid(object value)
        {
            string phoneNumber = (string)value;
            string pattern = @"\d{4}-\d{6}";
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);

            if (rgx.IsMatch(phoneNumber))
                return true;
            else
                return false;


        }
    }
}