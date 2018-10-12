using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

using System.Diagnostics;
using System.Data.OleDb;

namespace cycleCalculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string url = System.Environment.CurrentDirectory;

        static DataTable machineDt = GetDataFromExcelByConn("data\\gcxxb.xls");
        Dictionary<string, machine> machineDic = getMachineListFromTable(machineDt);

        static DataTable materialDt = GetDataFromExcelByConn("data\\clxxb.xls");
        Dictionary<string, material> materialDic = getMaterialListFromTable(materialDt);


        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> machineNames = new List<string>();
            foreach(var m in machineDic)
            {
                machineNames.Add(m.Key);
            }
            comboBox1.DataSource = machineNames;

            List<string> materialNames = new List<string>();
            foreach(var m in materialDic)
            {
                materialNames.Add(m.Key);
            }
            comboBox2.DataSource = materialNames;

            //string[] machineNames = machineDic.Keys;
            //for (int i = 1; i < dt.Rows.Count; i++)
            //{
            //    for (int j = 1; j < dt.Columns.Count; j++)
            //    {
            //        Console.Write(dt.Rows[i][j]);
            //        Console.Write(" ");
            //    }
            //}
        }

        //由表格数据得到机器字典
        public static Dictionary<string, machine> getMachineListFromTable(DataTable dt)
        {
            Dictionary<string, machine> machineDic = new Dictionary<string, machine>();
            
            for(int i = 1; i < dt.Rows.Count; i++)
            {
                machine m = new machine();
                m.id = (string)dt.Rows[i][0];
                m.count = System.Convert.ToDouble(dt.Rows[i][1]);
                m.diameter = System.Convert.ToDouble(dt.Rows[i][2]);
                m.maxDistance = System.Convert.ToDouble(dt.Rows[i][3]);
                m.maxVelocity = System.Convert.ToDouble(dt.Rows[i][4]);
                m.cost = System.Convert.ToDouble(dt.Rows[i][5]);

                machineDic.Add(m.id, m);
            }

            return machineDic;
        }

        //由表格得到材料字典
        public static Dictionary<string, material> getMaterialListFromTable(DataTable dt)
        {
            Dictionary<string, material> materialDic = new Dictionary<string, material>();

            for (int i = 1; i < dt.Rows.Count; i++)
            {
                material m = new material();
                m.no = System.Convert.ToInt32(dt.Rows[i][0]);
                m.id = (string)dt.Rows[i][1];
                m.T0 = System.Convert.ToDouble(dt.Rows[i][2]);
                m.Td = System.Convert.ToDouble(dt.Rows[i][3]);
                m.t = System.Convert.ToDouble(dt.Rows[i][4]);
                m.Tw = System.Convert.ToDouble(dt.Rows[i][5]);
                m.Tc = System.Convert.ToDouble(dt.Rows[i][6]);
                m.Tr = System.Convert.ToDouble(dt.Rows[i][7]);
                m.p = System.Convert.ToDouble(dt.Rows[i][8]);
                m.a = System.Convert.ToDouble(dt.Rows[i][9]);

                materialDic.Add(m.id, m);
            }

            return materialDic;
        }



        public static DataTable GetDataFromExcelByConn(string filePath)
        {
            bool hasTitle = false;
            //var filePath = "C:\\Users\\yanglu\\Desktop\\gcxxb.xls";
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType)) return null;

            using (DataSet ds = new DataSet())
            {
                string strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                                "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                                "data source={3};",
                                (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                string strCom = " SELECT * FROM [Sheet1$]";
                using (OleDbConnection myConn = new OleDbConnection(strCon))
                using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn))
                {
                    myConn.Open();
                    myCommand.Fill(ds);
                }
                if (ds == null || ds.Tables.Count <= 0) return null;
                return ds.Tables[0];
            }
        }

        void getResult(double input1, double input2, double input3)
        {
            string machineName = comboBox1.Text;
            machine mac = new machine();
            machineDic.TryGetValue(machineName, out mac);
            string materialName = comboBox2.Text;
            material mat = new material();
            materialDic.TryGetValue(materialName, out mat);
        }

        public class machine
        {
            public string id { get; set; }
            public double count { get; set; }
            public double diameter { get; set; }
            public double maxDistance { get; set; }
            public double maxVelocity { get; set; }
            public double cost { get; set; }
        }

        public class material
        {
            public int no { get; set; }
            public string id { get; set; }
            public double T0 { get; set; }
            public double Td { get; set; }
            public double t { get; set; }
            public double Tw { get; set; }
            public double Tc { get; set; }
            public double Tr { get; set; }
            public double p { get; set; }
            public double a { get; set; }
        }

        //清除输出数据
        public void clearOutTextBox()
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";

        }

        //获取Ds值
        public double getDs(double m, double p, double D)
        {
            double result = 0.0;
            result = 4000 * m / (p * Math.PI * D * D);
            result = setTwo(result);
            return result;
        }

        //保留两位小数
        public double setTwo(double d)
        {
            d = (double)Math.Round(d * 100) / 100;
            return d;
        }

        //获取T1值
        public double getT1(double s, double a, double t0, double tr, double tw)
        {
            double t1 = 0.0;
            double d = 0.3 + s / 2;
            t1 = ((d * d) / (9.87 * a)) * Math.Log(1.273 * (t0 - tw) / (tr - tw));
            t1 = setTwo(t1);
            return t1;
        }

        //获取T2值
        public double getT2(double s, double a, double t0, double tc, double tw)
        {
            double t2 = 0.0;
            t2 = (s * s / (9.87 * a)) * Math.Log(1.103 * (t0 - tw) / (tc - tw));
            t2 = setTwo(t2);
            return t2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start("C:\\Users\\yanglu\\Desktop\\gcxxb.xls");
            //System.Diagnostics.Process.Start("C:\\Users\\yanglu\\Desktop\\gcxxb.xls");
            //System.Diagnostics.Process.Start("C:\\Users\\yanglu\\Desktop\\hhh.xlsx");

            //Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            //app.Visible = true;
            //app.Workbooks.Add("C:\\Users\\yanglu\\Desktop\\hhh.xlsx");
            //Workbook book = app.Workbooks.Open("D:\\Test.xlsx ");
            //Shell("c:\windows\explorer.exe d:");

            //System.Diagnostics.Process.Start(@"C:\Users\yanglu\Desktop\gcxxb.xls");

            Process.Start("C:\\Users\\yanglu\\Desktop\\hhh.xlsx");

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string input1 = textBox1.Text;
            string input2 = textBox2.Text;
            string input3 = textBox3.Text;
            if(input1 == "" || input2 == "" || input3 == "")
            {
                MessageBox.Show("请先输入数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            string machineName = comboBox1.Text;
            machine mac = new machine();
            machineDic.TryGetValue(machineName, out mac);
            string materialName = comboBox2.Text;
            material mat = new material();
            materialDic.TryGetValue(materialName, out mat);
            textBox4.Text = mat.T0.ToString();
            textBox5.Text = mat.Td.ToString();
            textBox6.Text = mat.t.ToString();
            textBox7.Text = mat.Tw.ToString();
            double Ds = getDs(System.Convert.ToDouble(input1), mat.p, mac.diameter);
            textBox8.Text = Ds.ToString();
            textBox9.Text = setTwo(Ds / 5).ToString();
            double o7 = (Ds + 4) / (mac.maxVelocity * 0.3 * 0.5);
            textBox10.Text = o7.ToString();
            double t1 = getT1(System.Convert.ToDouble(input2), mat.a, mat.T0, mat.Tr, mat.Tw);
            textBox11.Text = t1.ToString();
            double t2 = getT2(System.Convert.ToDouble(input2), mat.a, mat.T0, mat.Tc, mat.Tw);
            textBox12.Text = t2.ToString();
            textBox13.Text = setTwo(o7 + t1 + t2 + System.Convert.ToDouble(input3)).ToString();

            if (Ds > 0.6 * mac.diameter)
            {
                clearOutTextBox();
                MessageBox.Show("选择的机台熔胶量不够", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                {
                    ctr.Text = "";
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            machineDt = GetDataFromExcelByConn("data\\gcxxb.xls");
            machineDic = getMachineListFromTable(machineDt);

            materialDt = GetDataFromExcelByConn("data\\clxxb.xls");
            materialDic = getMaterialListFromTable(materialDt);
        }
    }
}
