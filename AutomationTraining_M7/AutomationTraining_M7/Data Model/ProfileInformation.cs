using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Model
{
    class ProfileInformation
    {
        public string name { get; set; }
        public string lastname { get; set; }
        public string middlename { get; set; }
        public int age { get; set; }
        public string email { get; set; }
        public string mobile_phone { get; set; }
        public maritalStatus marital_status { get; set; }

        public enum maritalStatus
        {
            Married = 0,
            Widowed = 1,
            Separated = 2,
            Divorced = 3,
            Single = 4
        }

    }
}
