using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using DataAccessUtilExcel.Data;
using DataAccessUtilExcel.Data.Excel;

namespace WebTest
{

    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            IDataAccessUtil excel = new DataAccessUtil(@"E:\BaseDatos\example.xlsx");

            excel.SetSheetName("Hoja1");
            excel.SetColumns(new string[] { "ITEM", "APELLIDO PATERNO", "APELLIDO MATERNO", "TIPO DE ENTIDAD", "ENTIDAD", "CARGO" });

            var data = excel.GetDataTable();

            gvResult.DataSource = data;
            gvResult.DataBind();

            //excel.SetSheetName("Hoja1");
            //excel.SetColumns(new string[] { "ITEM", "TIPO DE ENTIDAD", "ENTIDAD", "CARGO", "APELLIDO PATERNO", "NOMBRE1" });
            //excel.SetValues(new string[] { "155", "GOBIERNO CENTRAL", "MINISTERIO DE COMERCIO EXTERIOR Y TURISMO", "VICE MINISTRO DE COMERCIO EXTERIOR", "POSADA", "VLADO" });
            //excel.Insert();

        }
    }
}