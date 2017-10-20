using System;
using System.Threading;
using System.Windows.Forms;
using System.Data;
using System.Drawing;

using ad2csv.Db.LDAP;

namespace ad2csv_test
{
    public partial class Form1 : Form
    {
        bool isRunning = false;

        public Form1()
        {
            InitializeComponent();
            tbDomain.Text = "domain.bc.ca";
            tbUsr.Text = "administrator";
            tbPwd.Text = "";
            SetText(textBox1, this.BackColor, "Ready");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnExtract_Click(object sender, EventArgs e)
        {
            if(!isRunning) {
                isRunning = true;
                SetText(textBox1, Color.Red, "Running");
                Thread oThread = new Thread(RunExtract);
                oThread.Start();
            } else {
                System.Console.WriteLine("Already running!");
            }
        }

        protected void RunExtract() {
            ADtoDataTable oExtractor = new ADtoDataTable();
            DataSet oDs = oExtractor.Ad2DataTable(tbDomain.Text, tbUsr.Text, tbPwd.Text);
            SetGridSource(dataGridView1, oDs.Tables["AD_Users"]);
            SetGridSource(dataGridView2, oDs.Tables["AD_Groups"]);
            SetGridSource(dataGridView3, oDs.Tables["AD_Units"]);
            SetText(textBox1, this.BackColor, "Ready");
            isRunning = false;;
        }


        delegate void SetDataSourceCallback(DataGridView grid, DataTable oDt);
        private void SetGridSource(DataGridView grid, DataTable oDt)
        {
          // Compare the threadID of the calling thread to the threadID of the creating thread.
          //        true if different
          if (grid.InvokeRequired) { 
            SetDataSourceCallback d = new SetDataSourceCallback(SetGridSource);
            this.Invoke(d, new object[] { grid, oDt });
          } else {
            grid.DataSource = oDt;
          }
        }

        delegate void SetTextCallback(TextBox tb, Color backClr, String msg);
        private void SetText(TextBox tb, Color backClr, String msg)
        {
          // Compare the threadID of the calling thread to the threadID of the creating thread.
          //        true if different
          if (tb.InvokeRequired) { 
            SetTextCallback d = new SetTextCallback(SetText);
            this.Invoke(d, new object[] { tb, backClr, msg });
          } else {
            tb.Text = msg;
            tb.BackColor = backClr;
          }
        }
    }
}
