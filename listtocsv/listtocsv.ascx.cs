namespace ArenaWeb.UserControls.Custom.WVC
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Linq;
    using System.Text;
    using System.Web;
    using System.Web.UI;
    using System.Web.UI.WebControls;
    using Arena.Portal;
    using Arena.Core;
    using Arena.DataLayer;
    using Arena.List;

    public partial class ListToCSV : PortalControl
    {
        #region Module Settings

        // Module Settings
        [Setting("Token Variable", "The query string token provided on the querystring", true)]
        public string TokenVariableSetting { get { return Setting("TokenVariable", "", true); } }

        [Setting("Token Value", "The query string token value provided on the querystring", true)]
        public string TokenValueSetting { get { return Setting("TokenValue", "", true); } }

        [Setting("Reverse DNS Calling Server", "The domain that is allowed to call this page", true)]
        public string RDNSCallerValueSetting { get { return Setting("RDNSCallerValue", "", true); } }

        #endregion

        // ArenaWeb.UserControls.List.ListReportView
        private int _reportId = -1;

        // ArenaWeb.UserControls.List.ListReportView
        private ListReport rpt;

        private string output;

        protected void validateToken(string value)
        {
            if (value != TokenValueSetting)
            {
                throw new ModuleException(base.CurrentPortalPage, base.CurrentModule, "The token value doesn't match.");
            }
        }

        

        protected void Page_Init(object sender, EventArgs e)
        {
            if (!int.TryParse(base.Request.QueryString["reportid"].ToString(), out this._reportId)) { }
            if (base.Request.QueryString[TokenVariableSetting].ToString().Length != 0)
            {
                this.validateToken(Request.QueryString[TokenVariableSetting]);
            }
            /*if (!System.Net.Dns.GetHostEntry(base.Request.UserHostAddress).HostName.Contains(RDNSCallerValueSetting))
            {
                throw new ModuleException(base.CurrentPortalPage, base.CurrentModule, System.Net.Dns.GetHostEntry(base.Request.UserHostAddress).HostName + " is not allowed to call this page.");
            }*/
            
            this.rpt = new ListReport(this._reportId);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            SqlDataReader reader = null;
            DataTable listDT = null;
            reader = new Arena.DataLayer.Organization.OrganizationData().ExecuteReader(this.rpt.Query);
            listDT = SqlReaderToDataTable(reader);

            Response.Clear();
            Response.ContentType="text/csv";

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < listDT.Columns.Count; i++)
            {
                sb.Append(Convert.ToChar(34) + listDT.Columns[i].ColumnName + Convert.ToChar(34) + ',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(Environment.NewLine);
            for (int j = 0; j < listDT.Rows.Count; j++)
            {
                for (int k = 0; k < listDT.Columns.Count; k++)
                {
                    sb.Append(Convert.ToChar(34) + listDT.Rows[j][k].ToString() + Convert.ToChar(34) + ',');
                }
                sb.Remove(sb.Length - 1, 1);
                sb.Append(Environment.NewLine);
            }
            Response.Write(sb.ToString());
            Response.End();
        }

        static private DataTable SqlReaderToDataTable(SqlDataReader rdr)
        {
            DataTable dt = new DataTable("hdc_customTable");
            DataRow row;
            int i;
            for (i = 0; i < rdr.FieldCount; i++)
            {
                dt.Columns.Add(rdr.GetName(i));
            }
            while (rdr.Read())
            {
                row = dt.NewRow();
                for (i = 0; i < rdr.FieldCount; i++)
                {
                    row[i] = rdr[i];
                }
                dt.Rows.Add(row);
            }
            return dt;	
        }
    }
}