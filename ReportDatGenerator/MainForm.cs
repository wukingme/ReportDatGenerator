using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportDatGenerator
{
    public partial class MainForm : Form
    {
        private string strDocumnetSavePath = string.Empty;
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.strGenerateDatFileName = $"{DateTime.Now:yyyyMMddHHmmss}.dat";
            this.txtDatFileName.Text = this.strGenerateDatFileName;
            this.strDocumnetSavePath = Path.Combine(Application.ExecutablePath, $"{DateTime.Now:yyyyMMdd}", this.strGenerateDatFileName);
        }

        /// <summary>
        /// 当前选择的Excel文件
        /// </summary>
        private string strCurrentDataFile = string.Empty;

        /// <summary>
        /// 需要生成的Dat文件名
        /// </summary>
        private string strGenerateDatFileName = string.Empty;
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(this.openFileDialog1.FileName))
                    this.strCurrentDataFile = this.openFileDialog1.FileName;
            }
        }

        /// <summary>
        /// 生成Dat文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strCurrentDataFile + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter myCommand = null;
            string strExcel = string.Format("select * from [{0}$]", "Sheet1");
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            myCommand.Fill(dt);
            StringBuilder sbdDatData = new StringBuilder();
            sbdDatData.AppendLine($"For M5|Adr     1|TO  {this.strCurrentDataFile}               |                      |                      |                      | ");
            sbdDatData.AppendLine($"For M5|Adr     2|TO  Start-Line      aBFFB     4|                      |                      |                      | ");
            sbdDatData.AppendLine($"For M5|Adr     3|KD1       J1                  4|                      |                      |Z         0.00000 m   | ");
            if (dt.Rows.Count > 0)
            {
                int rowIndex = 4, dataIndex = 0;
                double timeSpan = 15;
                double rb1 = 0.00000f, rf2_1 = 0.00000f, rf2_2 = 0.00000f, rb1_2 = 0.00000f,
                    z_Yzd = 0.00000f, //已知点
                    z_Wzd = 0.00000f;//未知点
                foreach (DataRow row in dt.Rows)
                {
                    dataIndex = dt.Rows.IndexOf(row);
                    if (dataIndex == 0) continue;//跳过标题行
                    rowIndex = rowIndex + dataIndex - 1;
                    z_Wzd = Convert.ToDouble(row[3]);
                    if (dataIndex % 2 == 0)
                    {
                        sbdDatData.AppendLine($"For M5|Adr     {dataIndex}|KD1       {row[0]}      {DateTime.Now.AddMilliseconds(timeSpan):HH:ss:iii}   4 |Rb        {rb1} m   |HD         10.253 m   |                      | ");
                        sbdDatData.AppendLine($"For M5|Adr     {dataIndex++}|KD1       {row[0]}      {DateTime.Now.AddMilliseconds(timeSpan):HH:ss:iii}   4 |Rf        {rb1} m   |HD         10.253 m   |                      | ");
                        sbdDatData.AppendLine($"For M5|Adr     {dataIndex++}|KD1       {row[0]}      {DateTime.Now.AddMilliseconds(timeSpan):HH:ss:iii}   4 |Rf        {rb1} m   |HD         10.253 m   |                      | ");
                        sbdDatData.AppendLine($"For M5|Adr     {dataIndex++}|KD1       {row[0]}      {DateTime.Now.AddMilliseconds(timeSpan):HH:ss:iii}   4 |Rb        {rb1} m   |HD         10.253 m   |                      | ");
                    }
                    sbdDatData.AppendLine($"For M5|Adr     {dataIndex++}|KD1       {row[0]}      {DateTime.Now.AddMilliseconds(timeSpan):HH:ss:iii}   4 |Rb        {rb1} m   |HD         10.253 m   |                      | ");
                }
            }
        }
    }
}
