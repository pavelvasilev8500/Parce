using System;
using System.IO;
using System.Text;

namespace Parce
{
    class ElevatorManager
    {
        private string sheetName { get; set; }
        private string[] Args { get; set; }
        public ElevatorManager(string[] args)
        {
            Args = args;
            ExecelParce exel = new ExecelParce(Args);
            sheetName = Validation.GetSheetName();
            while (!exel.CheckBookPath(sheetName))
            {
                Console.WriteLine("Неверное название книги!");
                sheetName = Validation.GetSheetName();
            }
            Menu();
            while (true)
            {
                Console.Write("Выбирите действие: ");
                var key = Console.ReadKey().Key;
                Console.WriteLine();
                switch (key)
                {
                    case ConsoleKey.NumPad1:
                        Set4N();
                        break;
                    case ConsoleKey.NumPad2:
                        Set4S();
                        break;
                    case ConsoleKey.C:
                        Clear();
                        break;
                    case ConsoleKey.M:
                        Menu();
                        break;
                    case ConsoleKey.Q:
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
            }
        }

        private async void Set4N()
        {
            var currentDate = new CurrentDate();
            var PATH = $"{Environment.CurrentDirectory}\\PH109{currentDate._year}{currentDate._month}.31";
            var resultData = new ResultData();
            var headerData = new HeaderData();
            var mainData = resultData.GetElevators(Args, sheetName, false);
            Console.WriteLine(headerData.Set4N(mainData.Item2));
            foreach (var item in mainData.Item1)
                Console.WriteLine($"{item._fOrgCode}" +
                                  $"{item._sOrgCode};" +
                                  $"{item._api};" +
                                  $"{item._street};" +
                                  $"{item._home};" +
                                  $"{item._date};" +
                                  $"{item._serviceCode};" +
                                  $"{item._serviceName};" +
                                  $"{item._counterId};" +
                                  $"{item._firstCounterValue};" +
                                  $"{item._lastCounterValue};" +
                                  $"{item._serviceScope};" +
                                  $"{item._ratio};" +
                                  $"{item._rate};" +
                                  $"{item._summ};" +
                                  $"{item._rateACT};" +
                                  $"{item._ACTsumm};" +
                                  $"{item._summWhithACT};");
            using (StreamWriter writer = new StreamWriter(PATH, false, Encoding.GetEncoding(1251)))
            {
                await writer.WriteLineAsync(headerData.Set4N(mainData.Item2));
                foreach(var item in mainData.Item1)
                    await writer.WriteLineAsync($"{item._fOrgCode}" +
                                  $"{item._sOrgCode};" +
                                  $"{item._api};" +
                                  $"{item._street};" +
                                  $"{item._home};" +
                                  $"{item._date};" +
                                  $"{item._serviceCode};" +
                                  $"{item._serviceName};" +
                                  $"{item._counterId};" +
                                  $"{item._firstCounterValue};" +
                                  $"{item._lastCounterValue};" +
                                  $"{item._serviceScope};" +
                                  $"{item._ratio};" +
                                  $"{item._rate};" +
                                  $"{item._summ};" +
                                  $"{item._rateACT};" +
                                  $"{item._ACTsumm};" +
                                  $"{item._summWhithACT};");
            }
        }

        private async void Set4S()
        {
            var currentDate = new CurrentDate();
            var PATH = $"{Environment.CurrentDirectory}\\PH109{currentDate._year}{currentDate._month}.33";
            var resultData = new ResultData();
            var headerData = new HeaderData();
            var mainData = resultData.GetElevators(Args, sheetName, true);
            Console.WriteLine(headerData.Set4S(mainData.Item2));
            foreach (var item in mainData.Item1)
                Console.WriteLine($"{item._fOrgCode}" +
                                  $"{item._sOrgCode};" +
                                  $"{item._api};" +
                                  $"{item._street};" +
                                  $"{item._home};" +
                                  $"{item._date};" +
                                  $"{item._serviceCode};" +
                                  $"{item._serviceName};" +
                                  $"{item._counterId};" +
                                  $"{item._firstCounterValue};" +
                                  $"{item._lastCounterValue};" +
                                  $"{item._serviceScope};" +
                                  $"{item._ratio};" +
                                  $"{item._rate};" +
                                  $"{item._summ};" +
                                  $"{item._rateACT};" +
                                  $"{item._ACTsumm};" +
                                  $"{item._summWhithACT};");
            using (StreamWriter writer = new StreamWriter(PATH, false, Encoding.GetEncoding(1251)))
            {
                await writer.WriteLineAsync(headerData.Set4S(mainData.Item2));
                foreach (var item in mainData.Item1)
                    await writer.WriteLineAsync($"{item._fOrgCode}" +
                                  $"{item._sOrgCode};" +
                                  $"{item._api};" +
                                  $"{item._street};" +
                                  $"{item._home};" +
                                  $"{item._date};" +
                                  $"{item._serviceCode};" +
                                  $"{item._serviceName};" +
                                  $"{item._counterId};" +
                                  $"{item._firstCounterValue};" +
                                  $"{item._lastCounterValue};" +
                                  $"{item._serviceScope};" +
                                  $"{item._ratio};" +
                                  $"{item._rate};" +
                                  $"{item._summ};" +
                                  $"{item._rateACT};" +
                                  $"{item._ACTsumm};" +
                                  $"{item._summWhithACT};");
            }
        }

        private void Clear()
        {
            Console.Clear();
        }

        private void Menu()
        {
            Console.Write(@"+---+-------------------------------+" + "\n" +
                          @"| 1 | Создать список для Новобелицы |" + "\n" +
                          @"+---+-------------------------------|" + "\n" +
                          @"| 2 | Создать список для Советского |" + "\n" +
                          @"+---+-------------------------------|" + "\n" +
                          @"| c | Очистить                      |" + "\n" +
                          @"+---+-------------------------------|" + "\n" +
                          @"| m | Меню                          |" + "\n" +
                          @"+---+-------------------------------|" + "\n" +
                          @"| q | Выход                         |" + "\n" +
                          @"+---+-------------------------------+" + "\n");
        }
    }
}
