using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelServices.EuroStrategy;
//using System.Web.Services.Protocols;

namespace ExcelServices
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelServiceSoapClient es = new ExcelServiceSoapClient();
            Status[] outStatus;
            RangeCoordinates rangeCoordinates = new RangeCoordinates();
            string sheetName = "Hoja1";
            string targetWorkbookPath = "http://eurostrategy.sharepoint.com/Documentos%20compartidos/Ejemplo%20Excel%20externo%2010%20-%20oData%20feed%20PLSQL.xlsx";
            //es..Credentials = System.Net.CredentialCache.DefaultCredentials;
                string sessionId = es.OpenWorkbook(targetWorkbookPath, "en-US", "en-US", out outStatus);
                es.CloseWorkbook(sessionId);
        }
    }
}
