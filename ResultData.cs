using System;
using System.Collections.Generic;
using System.Threading;

namespace Parce
{
    class ResultData
    {

        private List<ElevatorModel> ele;
        private double allsumm { get; set; } = 0;

        public (List<ElevatorModel>, double) GetElevators(string[] args, string sheetName, bool n)
        {
            ExecelParce exel = new ExecelParce(args);
            var currentDate = new CurrentDate();
            var dataList = new List<string>();
            int elevatorsCount = Validation.GetAElevatorsCount();
            int firsPosition = Validation.GetFirstPosition();
            ele = new List<ElevatorModel>();
            int count = elevatorsCount + firsPosition;
            Console.Write("Подготовка задачи... ");
            using (var progress = new ProgressBar())
            {
                for (int i = firsPosition; i < count; i++)
                {
                    dataList = exel.GetCells(i, new List<int>() { 2, 3, 4, 23, 100 }, sheetName);
                    if (!string.IsNullOrEmpty(dataList[0]))
                    {
                        try
                        {
                            if (dataList[2].ToString().Length < 2)
                            {
                                dataList[2] = "0" + dataList[2];
                            }
                            else
                            {
                                dataList[2] = dataList[2];
                            }
                        }
                        catch (Exception)
                        { }
                        var sh = dataList[1].Split(new char[] { ',' });
                        dataList[4] = sh[1].Trim().Split(new char[] { 'п' })[0].TrimEnd().Replace(" ", "-").Replace(".", "").ToUpper();
                        ele.Add(new ElevatorModel
                        {
                            _sOrgCode = dataList[2],
                            _street = dataList[0],
                            _home = dataList[4],
                            _date = $"01.{currentDate._month}.{currentDate.year}",
                            _summ = dataList[3],
                            _summWhithACT = dataList[3]
                        });
                        try
                        {
                            allsumm = allsumm + double.Parse(dataList[3]);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"{e} - {dataList[3]}");
                        }
                    }
                    progress.Report(((double)i - (firsPosition - 1)) / elevatorsCount);
                    Thread.Sleep(20);
                }
            }
            Console.WriteLine("Готово.");
            return (ele, allsumm);
        }
    }
}
