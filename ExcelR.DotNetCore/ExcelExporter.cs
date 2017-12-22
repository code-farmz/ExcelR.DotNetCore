using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelR.DotNetCore
{
  public  class ExcelExporter
    {
        private IWorkbook _workbook;

        public ExcelExporter()
        {
            _workbook = GetWorkbook();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IWorkbook GetWorkbook(string defaultSheetName = "Sheet1")
        {
            if (_workbook != null)
                return _workbook;
            _workbook = new XSSFWorkbook();
            _workbook.CreateSheet(defaultSheetName);
            return _workbook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ISheet GetWorkSheet(string name = "Sheet1")
        {
            if (_workbook == null)
                _workbook = new XSSFWorkbook();
            return _workbook.GetSheet(name) ?? _workbook.CreateSheet(name);
        }


    }
}
