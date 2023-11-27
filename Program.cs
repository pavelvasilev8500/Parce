using System;
using System.Threading;

namespace Parce
{
    static class Program
    {
        static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            new ElevatorManager(args);
            Console.ReadKey();
        }
    }
}
