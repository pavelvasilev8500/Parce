using System;

namespace Parce
{
    public static class Validation
    {
        private static bool result { get; set; }

        public static string GetSheetName()
        {
            Console.Write("Введите название книги:");
            return Console.ReadLine();
        }

        public static int GetFirstPosition()
        {
            int fp;
            Console.Write("Введите первую позицию:");
            result = int.TryParse(Console.ReadLine(), out fp);
            if (fp <= 0)
            {
                result = false;
            }
            while (!result)
            {
                Console.WriteLine("Неверный индекс!");
                Console.Write("Введите первую позицию: ");
                result = int.TryParse(Console.ReadLine(), out fp);
                if (fp < 0)
                {
                    result = false;
                }
            }
            return fp;
        }

        public static int GetAElevatorsCount()
        {
            int elevatorsCount;
            Console.Write("Введите количество лифтов:");
            result = int.TryParse(Console.ReadLine(), out elevatorsCount);
            if (elevatorsCount <= 0)
            {
                result = false;
            }
            while (!result)
            {
                Console.WriteLine("Неверное количество!");
                Console.Write("Введите количество лифтов: ");
                result = int.TryParse(Console.ReadLine(), out elevatorsCount);
                if (elevatorsCount < 0)
                {
                    result = false;
                }
            }
            return elevatorsCount;
        }
    }
}
