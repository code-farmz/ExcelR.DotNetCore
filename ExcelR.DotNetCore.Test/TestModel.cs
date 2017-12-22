using ExcelR.DotNetCore.Attributes;
using System;

namespace ExcelRTest
{
    public class TestModel
    {
        [ExcelRProp(Name = "First Name")]
        public string FirstName { get; set; }


        [ExcelRProp(ColTextColor = "Red", Name = "Last Name")]
        public string LastName { get; set; }

        [ExcelRProp(SkipExport = true)]
        public bool IsMale { get; set; }

        [ExcelRProp(HeadTextColor = "Blue" ,Name = "Date Of Birth")]
        public DateTime? Dob { get; set; }

    }

}
