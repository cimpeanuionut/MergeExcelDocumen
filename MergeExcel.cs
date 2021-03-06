﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace MergeExcelDocument
{
    public class MergeExcel
    {
        Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();

        //initialize the object of saved target
        Excel.Workbook bookDest = null;
        Excel.Worksheet sheetDest = null;

        //initialize the object of read data
        Excel.Workbook bookSource = null;
        Excel.Worksheet sheetSource = null;
        string[] _sourceFiles = null;
        string _destFile = string.Empty;
        string _columnEnd = string.Empty;
        int _headerRowCount = 0;
        int _currentRowCount = 0;

        public MergeExcel(string[] sourceFiles, string destFile, string columnEnd, int headerRowCount)
        {

            //Use class Missing case to indicate the missing value. e.g. when you call the method that has default parameter value
            bookDest = (Excel.WorkbookClass)app.Workbooks.Add(Missing.Value);
            sheetDest = bookDest.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Excel.Worksheet;
            sheetDest.Name = "Data";
            _sourceFiles = sourceFiles;
            _destFile = destFile;
            _columnEnd = columnEnd;
            _headerRowCount = headerRowCount;
        }

        //open worksheet
        void OpenBook(string fileName)
        {
            bookSource = app.Workbooks._Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            sheetSource = bookSource.Worksheets[1] as Excel.Worksheet;
        }

        //close worksheet
        void CloseBook()
        {
            bookSource.Close(false, Missing.Value, Missing.Value);
        }

        //copy table header
        void CopyHeader()
        {
            Excel.Range range = sheetSource.get_Range("B1", _columnEnd + _headerRowCount.ToString());
            range.Copy(sheetDest.get_Range("B1", Missing.Value));
            _currentRowCount += _headerRowCount;
        }

        //copy data
        void CopyData()
        {
            int sheetRowCount = sheetSource.UsedRange.Rows.Count;
            Excel.Range range = sheetSource.get_Range(string.Format("B{0}", _headerRowCount), _columnEnd + sheetRowCount.ToString());
            range.Copy(sheetDest.get_Range(string.Format("B{0}", _currentRowCount), Missing.Value));
            _currentRowCount += range.Rows.Count;
        }

        //save the result
        void Save()
        {
            bookDest.Saved = true;
            bookDest.SaveCopyAs(_destFile);
        }

        //exit the process
        void Quit()
        {
            app.Quit();
        }
        void DoMerge()
        {

            //declare variate bool to judge if copy table header
            bool b = false;
            foreach (string strFile in _sourceFiles)
            {
                OpenBook(strFile);
                if (b == false)
                {
                    CopyHeader();
                    b = true;
                }
                CopyData();
                CloseBook();
            }
            Save();
            Quit();
        }

        public static void DoMerge(string[] sourceFiles, string destFile, string columnEnd, int headerRowCount)
        {
            new MergeExcel(sourceFiles, destFile, columnEnd, headerRowCount).DoMerge();
        }

    }
}
