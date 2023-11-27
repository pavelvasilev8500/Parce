using System;

namespace Parce
{
    class HeaderData
    {
        private int noact { get; set; }

        CurrentDate currentDate = new CurrentDate();

        public HeaderData()
        {
            switch (currentDate.month)
            {
                case 0:
                    noact = 12;
                    break;
                case 1:
                    noact = 1;
                    break;
                case 2:
                    noact = 2;
                    break;
                case 3:
                    noact = 3;
                    break;
                case 4:
                    noact = 4;
                    break;
                case 5:
                    noact = 5;
                    break;
                case 6:
                    noact = 6;
                    break;
                case 7:
                    noact = 7;
                    break;
                case 8:
                    noact = 8;
                    break;
                case 9:
                    noact = 9;
                    break;
                case 10:
                    noact = 10;
                    break;
                case 11:
                    noact = 11;
                    break;
            }
        }
        public string Set4N(double allsum)
        {
            return $"410079672;\"Советское_Лифт\";491330177;КЖРЭУП \"Новобелицкое Гомель\";29.06.2022;1070;{currentDate.day}.{currentDate._month}.{currentDate.year};{noact};{allsum};0;{allsum}";
        }

        public string Set4S(double allsum)
        {
            return $"410079672;\"Советское_Лифт\";400079672;КЖРЭУП \"Советское\";15.08.2022;1133;{currentDate.day}.{currentDate._month}.{currentDate.year};{noact+1};{allsum};0;{allsum}";
        }

    }
}
