using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataAccessUtilExcel.Data
{
    public interface IDataAccessUtil
    {
        void SetSheetName(string sheetName);
        void SetColumns(string[] columns);
        void SetValues(string[] values);
        void Insert();
        void Insert(string[] values);
        void Insert(string[] columns, string[] values);
        void Insert(string sheet, string[] columns, string[] values);
        System.Data.DataTable GetDataTable(string query);
        System.Data.DataTable GetDataTable();
    }
}
