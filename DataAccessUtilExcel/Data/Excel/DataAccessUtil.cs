using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace DataAccessUtilExcel.Data.Excel
{
    using Common;

    public class DataAccessUtil : IDataAccessUtil
    {
        private readonly string PathFile;
        private string SheetName;
        private string[] Columns;
        private string[] Values;

        public DataAccessUtil(string pathFile)
        {
            this.PathFile = pathFile;
        }

        public DataAccessUtil(string pathFile, string sheetName)
        {
            this.PathFile = pathFile;
            this.SheetName = sheetName;
        }

        public void SetSheetName(string sheetName)
        {
            this.SheetName = sheetName;
        }

        public void SetColumns(string[] columns)
        {
            this.Columns = columns;
        }

        public void SetValues(string[] values)
        {
            this.Values = values;
        }

        protected IEnumerable<string> FormatColumns(string[] columns)
        {
            return columns.Select(c => string.Format("[{0}]", c));
        }

        protected IEnumerable<string> FormatValues(string[] values)
        {
            IList<string> paramsValue = new List<string>();

            for (int i = 0; i < values.Count(); i++)
                paramsValue.Add(string.Format("@Value{0}", i));

            return paramsValue; //values.Select(c => string.Format("@Value{0}", c));
        }

        protected string GetConnectionString()
        {
            return string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES;IMEX=0\"", this.PathFile);
        }

        protected IEnumerable<OleDbParameter> GetParameters()
        {
            IList<OleDbParameter> parameters = new List<OleDbParameter>();

            for (int i = 0; i < this.Values.Count(); i++)
			{
                parameters.Add(new OleDbParameter(string.Format("@Value{0}", i), this.Values[i]));
			}

            return parameters;
        }

        protected string GenerateInsertQuery()
        {
            if (string.IsNullOrEmpty(this.SheetName))
                throw new Exception("Error: sheet name is Null");

            if (this.Columns == null || !this.Columns.Any())
                throw new Exception("Error: columns is Null");

            var queryString = string.Empty;

            queryString = string.Format("INSERT INTO [{0}$] ({1}) VALUES ({2})", this.SheetName, string.Join(",", this.FormatColumns(this.Columns)), string.Join(",", this.FormatValues(this.Values)));

            return queryString;
        }

        protected string GenerateSelectQuery()
        {
            if (string.IsNullOrEmpty(this.SheetName))
                throw new Exception("Error: sheet name is Null");

            if (this.Columns == null || !this.Columns.Any())
                throw new Exception("Error: columns is Null");

            var queryString = string.Empty;
            
            queryString = string.Format("SELECT {1} FROM [{0}$]", this.SheetName, string.Join(",", this.FormatColumns(this.Columns)));

            return queryString;
        }

        protected int ExecuteNonQuery()
        {
            using (var cnx = new OleDbConnection(this.GetConnectionString()))
            {
                cnx.Open();
                using (var cmd = new OleDbCommand(this.GenerateInsertQuery(), cnx))
                {
                    cmd.Parameters.AddRange(GetParameters().ToArray());
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        public void Insert()
        {
            try
            {
                var result = ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Insert(string[] values)
        {
            try
            {
                this.SetValues(values);
                var result = ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Insert(string[] columns, string[] values)
        {
            try
            {
                this.SetColumns(columns);
                this.SetValues(values);
                var result = ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Insert(string sheetName, string[] columns, string[] values)
        {
            try
            {
                this.SetSheetName(sheetName);
                this.SetColumns(columns);
                this.SetValues(values);
                var result = ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetDataTable(string query)
        {
            using (var cnx = new OleDbConnection(this.GetConnectionString()))
            {
                cnx.Open();
                using (var cmd = new OleDbCommand(query, cnx))
                {
                    using(var da = new OleDbDataAdapter(cmd))
                    {
                        var dataSet = new DataSet();
                        da.Fill(dataSet);

                        return dataSet.Tables[0];
                    }
                }
            }
        }

        public DataTable GetDataTable()
        {
            using (var cnx = new OleDbConnection(this.GetConnectionString()))
            {
                cnx.Open();
                using (var cmd = new OleDbCommand(this.GenerateSelectQuery(), cnx))
                {
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        var dataSet = new DataSet();
                        da.Fill(dataSet);

                        return dataSet.Tables[0];
                    }
                }
            }
        }
    }
}
