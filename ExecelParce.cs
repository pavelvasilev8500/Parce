using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Parce
{
    class ExecelParce
    {
        private bool _checkPath { get; set; } = false;
        private string _sheetName { get; set; }
        //private string sheetName { get; set; }
        private string _value { get; set; } = "";
        private List<string> _valueList { get; set; }
        private string Path { get; set; }
        public ExecelParce(string[] path)
        {
            Path = String.Join(" ", path);
        }

        public bool CheckBookPath(string sheetName)
        {
            object rOnly = true;
            Application app = new Application();
            Workbooks workbooks = app.Workbooks;
            try
            {
                var workbook = workbooks.Open(Path, rOnly);
                dynamic ws;
                _sheetName = sheetName;
                if (string.IsNullOrEmpty(_sheetName))
                {
                    _sheetName = "1";
                    ws = workbook.Worksheets[int.Parse(_sheetName)];
                }
                else
                    ws = workbook.Worksheets[_sheetName];
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                    _checkPath = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                _checkPath = false;
            }
            finally
            {
                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
            return _checkPath;
        }

        public string GetCell(int row, int colum, string sheetName)
        {
            object rOnly = true;
            Application app = new Application();
            Workbooks workbooks = app.Workbooks;
            var workbook = workbooks.Open(Path, rOnly);
            if (workbook != null)
            {
                _sheetName = sheetName;
                if (string.IsNullOrEmpty(_sheetName))
                    _sheetName = "1";
                var ws = workbook.Worksheets[int.Parse(_sheetName)];
                _value = ws.UsedRange.Cells[row, colum].Value2.ToString();
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            if (workbooks != null)
            {
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
            return _value;
        }

        public List<string> GetCells(int row, List<int> colum, string sheetName)
        {
            object rOnly = true;
            Application app = new Application();
            Workbooks workbooks = app.Workbooks;
            var workbook = workbooks.Open(Path, rOnly);
            if (workbook != null)
            {
                _sheetName = sheetName;
                dynamic ws;
                if (string.IsNullOrEmpty(_sheetName))
                {
                    _sheetName = "2";
                    ws = workbook.Worksheets[int.Parse(_sheetName)];
                }
                else
                    ws = workbook.Worksheets[_sheetName];
                _valueList = new List<string>();
                foreach (var o in colum)
                    _valueList.Add((ws.UsedRange.Cells[row, o].Value2 ?? string.Empty).ToString());
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            if (workbooks != null)
            {
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
            return _valueList;
        }
    }
}
