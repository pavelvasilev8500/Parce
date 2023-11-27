using System;

namespace Parce
{
    class CurrentDate
    {
        public int day { get; set; }
        public int month { get; set; }
        public int year { get; set; }
        public string _month { get; set; }
        public string _year { get; set; }

        public CurrentDate()
        {
            DateTime now = DateTime.Now;
            month = now.Month;
            year = now.Year;
            _year = year.ToString().Substring(3);
            month = month - 1;
            if (month == 0)
            {
                month = 12;
                year = year - 1;
            }
            day = DateTime.DaysInMonth(year, month);
            if (month.ToString().Length < 2)
            {
                _month = "0" + month.ToString();
            }
            else
            {
                _month = month.ToString();
            }
        }
    }
}
