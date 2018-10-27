using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Net.Http;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using ICSharpCode.SharpZipLib.Zip;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        public class BankDataObj
        {
            public BankDataObj()
            {


            }

            public string CustomerCode { get; set; }
            public double Money { get; set; }
            public string Name { get; set; }
            public string ID { get; set; }
            public string City { get; set; }
            public string Province { get; set; }
            public string _InTime { get; set; }
            public DateTime InTime { get { return Convert.ToDateTime(_InTime); } set {; } }
        }

        public class ExceptObj
        {
            public ExceptObj()
            {

            }

            public string Name { get; set; }
            public string ID { get; set; }

            public string _ID { get { return !string.IsNullOrEmpty(ID) && ID.Length >= 16 ? ID.Replace(ID.Substring(10, 6), "******") : ""; } set {; } }
        }




        [DllImport("kernel32")]

        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        System.Timers.Timer pTimer = new System.Timers.Timer(50);//每隔5秒执行一次，没用winfrom自带的

        /// <summary>
        /// DataSetToList
        /// </summary>
        /// <typeparam name="T">转换类型</typeparam>
        /// <param name="dataSet">数据源</param>
        /// <param name="tableIndex">需要转换表的索引</param>
        /// <returns></returns>
        public List<T> DataSetToList<T>(DataSet dataSet, int tableIndex)
        {
            //确认参数有效
            if (dataSet == null || dataSet.Tables.Count <= 0 || tableIndex < 0)
            {
                return new List<T>();
            }

            System.Data.DataTable dt = dataSet.Tables[tableIndex];

            List<T> list = new List<T>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //创建泛型对象
                T _t = Activator.CreateInstance<T>();
                //获取对象所有属性
                PropertyInfo[] propertyInfo = _t.GetType().GetProperties().ToArray();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //按照属性顺序对应一一映射
                    if (dt.Rows[i][j] != DBNull.Value)
                    {
                        propertyInfo[j].SetValue(_t, dt.Rows[i][j], null);
                    }
                    else
                    {
                        propertyInfo[j].SetValue(_t, null, null);
                    }

                }
                list.Add(_t);
            }
            return list;
        }


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox5.Text = "200";  //默认总额度为200万
            textBox7.Text = "4000";  // 默认大于4000
            textBox8.Text = "200000";  //默认小于200000
            textBox9.Text = "58000";  //默认新案额度为58000
            //textBox2.Text= Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            textBox4.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            textBox3.Text = "dwquanziduan";

            pTimer.Elapsed += pTimer_Elapsed;//委托，要执行的方法
            pTimer.AutoReset = true;//获取该定时器自动执行
            pTimer.Enabled = false;//这个一定要写，要不然定时器不会执行的
            Control.CheckForIllegalCrossThreadCalls = false;

            //默认勾选所有选项
            for (int i = 0; i < 10; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }

        }

        private void pTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            pTimer.Enabled = false;

            string zipFilePath = GetOneFileByKeyWord(textBox4.Text, ".zip", textBox3.Text);

            if (!string.IsNullOrEmpty(zipFilePath))
            {
                //解压
                UnZipFile(zipFilePath, textBox4.Text);
            }

            string xlsFilePath = GetOneFileByKeyWord(textBox4.Text, ".xls", textBox3.Text);

            if (!string.IsNullOrEmpty(xlsFilePath))
            {
                button6.Text = "自动生成结果(已停止)";

                button2_Click(button2, e);
                button5_Click(button5, e);

                return;
            }


            pTimer.Enabled = true;
            return;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = folderBrowserDialog2.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                textBox4.Text = folderBrowserDialog2.SelectedPath;
                //MessageBox.Show(folderBrowserDialog1.SelectedPath);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog3 = new OpenFileDialog();     //显示选择文件对话框
            openFileDialog3.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);   //默认打开当前用户桌面路径
            openFileDialog3.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";  //只选取Excel文件
            openFileDialog3.FilterIndex = 2;
            openFileDialog3.RestoreDirectory = true;

            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = openFileDialog3.FileName;          //显示文件路径
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //ExcelToDS(textBox1.Text);
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("请选择导入的Excel");
                return;
            }

            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("请选择导出Excel的路径");
                return;
            }

            if (string.IsNullOrEmpty(textBox7.Text) || string.IsNullOrEmpty(textBox8.Text) || string.IsNullOrEmpty(textBox5.Text) || string.IsNullOrEmpty(textBox9.Text))
            {
                MessageBox.Show("金额或者额度不能为空！");
                return;
            }

            //DataSet data = new DataSet();

            //data = GetExcelTableByOleDB(textBox1.Text, "sheet1$", "select [客户代码] as 主持卡人代码,[余额(人民币)] from [sheet1$] where [城市调整] = '南宁' and [户籍城市] = '广西' and [余额(人民币)] > 4000 and [余额(人民币)] < 200000 and ([证件号码] LIKE '%1985%' or [证件号码] LIKE '%1986%' or [证件号码] LIKE '%1987%' or [证件号码] LIKE '%1988%' or [证件号码] LIKE '%1989%' or [证件号码] LIKE '%199%') order by [余额(人民币)] desc");

            //data = GetExcelTableByOleDB(textBox1.Text, "sheet1$", "select [持卡人姓名],[余额(人民币)],[客户代码],[证件号码],[城市调整],[户籍城市],[进案池时间] from [sheet1$] where [城市调整] = '南宁' and [户籍城市] = '广西' and [余额(人民币)] > 4000 and [余额(人民币)] < 200000 and ([证件号码] LIKE '%1985%' or [证件号码] LIKE '%1986%' or [证件号码] LIKE '%1987%' or [证件号码] LIKE '%1988%' or [证件号码] LIKE '%1989%' or [证件号码] LIKE '%199%') order by [余额(人民币)] desc");

            var data = GetExcelTableByOleDB(textBox1.Text, "sheet1$", "select [客户代码],[余额(人民币)],[持卡人姓名],[证件号码],[城市调整],[户籍城市],[进案池时间] from [sheet1$] where [城市调整] = '南宁' and [户籍城市] = '广西' and [余额(人民币)] > " + textBox7.Text + " and [余额(人民币)] < " + textBox8.Text + "  and [本金人民币]>0 order by [余额(人民币)] desc");

            var ExpData = GetExcelTableByOleDB(textBox10.Text, "sheet1$", "select * from [sheet1$]");


            if (data == null)
            {
                MessageBox.Show("有异常发生，无法生成结果，请联系开发人员！");
            }

            var bankdatalist = DataSetToList<BankDataObj>(data, 0);

            var ExpDataList = DataSetToList<ExceptObj>(ExpData, 0);

            //被排除清单
            List<BankDataObj> ExceptList = new List<BankDataObj>();


            foreach (var one in ExpDataList)
            {
                //如果有身份证号，则组合判断姓名及身份证
                if (!string.IsNullOrEmpty(one._ID))
                {
                    ExceptList.AddRange(bankdatalist.Where(r => r.Name.Equals(one.Name) && r.ID.Equals(one._ID)).ToList());
                }
                else
                {
                    //没有身份证号，则仅判断姓名
                    ExceptList.AddRange(bankdatalist.Where(r => r.Name.Equals(one.Name)).ToList());
                }


            }

            bankdatalist = bankdatalist.Except(ExceptList).ToList();

            List<BankDataObj> Level0 = new List<BankDataObj>();


            List<BankDataObj> ExpList = new List<BankDataObj>();
            List<BankDataObj> InlList = new List<BankDataObj>();

            //限制额度
            var LimitTotal = Convert.ToDouble(textBox5.Text) * 10000;
            double Total = 0;
            bool ContinueFlag = true;

            List<BankDataObj>[] AllLevel = new List<BankDataObj>[10];

            //如果选中年月优先
            if (checkBox2.CheckState == CheckState.Checked)
            {
                Level0 = bankdatalist.Where(x => (x.InTime.Year == dateTimePicker1.Value.Year && x.InTime.Month == dateTimePicker1.Value.Month)).OrderByDescending(r => r.Money).ToList();

                //bankdatalist = bankdatalist.Where(x => (x.InTime.Year != dateTimePicker1.Value.Year && x.InTime.Month != dateTimePicker1.Value.Month)).OrderByDescending(r => r.Money).ToList();

                //bankdatalist = bankdatalist.Where(x =>x._InTime.Substring(0, 4)!=dateTimePicker1.Value.Year.ToString()&& Convert.ToDateTime(x._InTime).ToString().Substring(4, 2) != dateTimePicker1.Value.Month.ToString()).OrderByDescending(r => r.Money).ToList();

                bankdatalist = bankdatalist.Where(x => new DateTime(x.InTime.Year, x.InTime.Month, 1) != new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, 1)).OrderByDescending(r => r.InTime).ToList();


                //90后 +4501
                var xLevel0 = Level0.Where(x => x.ID.Substring(6, 3).Equals("199") && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

                //90后 +非4501
                var xLevel1 = Level0.Where(x => x.ID.Substring(6, 3).Equals("199") && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();

                //85后+4501
                var xLevel2 = Level0.Where(x => (x.ID.Substring(6, 4).Equals("1985") || x.ID.Substring(6, 4).Equals("1986") || x.ID.Substring(6, 4).Equals("1987") || x.ID.Substring(6, 4).Equals("1988") || x.ID.Substring(6, 4).Equals("1989"))

                                                && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();
                //85后 +非4501
                var xLevel3 = Level0.Where(x => (x.ID.Substring(6, 4).Equals("1985") || x.ID.Substring(6, 4).Equals("1986") || x.ID.Substring(6, 4).Equals("1987") || x.ID.Substring(6, 4).Equals("1988") || x.ID.Substring(6, 4).Equals("1989"))

                                            && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();
                //80-84 +4501
                var xLevel4 = Level0.Where(x => (x.ID.Substring(6, 4).Equals("1985") || x.ID.Substring(6, 4).Equals("1986") || x.ID.Substring(6, 4).Equals("1987") || x.ID.Substring(6, 4).Equals("1988") || x.ID.Substring(6, 4).Equals("1989"))

                                                && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

                //80-84  +非4501
                var xLevel5 = Level0.Where(x => (x.ID.Substring(6, 4).Equals("1980") || x.ID.Substring(6, 4).Equals("1981") || x.ID.Substring(6, 4).Equals("1982") || x.ID.Substring(6, 4).Equals("1983") || x.ID.Substring(6, 4).Equals("1984"))

                                            && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();

                //70后 +4501
                var xLevel6 = Level0.Where(x => x.ID.Substring(6, 3).Equals("197") && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

                //70后 +非4501
                var xLevel7 = Level0.Where(x => x.ID.Substring(6, 3).Equals("197") && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();

                List<BankDataObj> xExpList = new List<BankDataObj>();
                List<BankDataObj> xInlList = new List<BankDataObj>();

                //限制额度
                var xLimitTotal = Convert.ToDouble(textBox9.Text);
                double xTotal = 0;

                List<BankDataObj>[] xAllLevel = new List<BankDataObj>[8];
                for (int i = 0; i < 8; i++)
                {
                    switch (i)
                    {
                        case 0: xAllLevel[i] = xLevel0; break;
                        case 1: xAllLevel[i] = xLevel1; break;
                        case 2: xAllLevel[i] = xLevel2; break;
                        case 3: xAllLevel[i] = xLevel3; break;
                        case 4: xAllLevel[i] = xLevel4; break;
                        case 5: xAllLevel[i] = xLevel5; break;
                        case 6: xAllLevel[i] = xLevel6; break;
                        case 7: xAllLevel[i] = xLevel7; break;
                        default: break;
                    }
                }

                //按优先级高到低(Level1~Level10)取数
                foreach (var one in xAllLevel)
                {
                    if (one != null && one.Count != 0)
                    {
                        xTotal = xTotal + one.Sum(x => x.Money);

                        if (xTotal < xLimitTotal)
                        {
                            xInlList.AddRange(one);
                            continue;
                        }
                        else
                        {
                            xExpList.AddRange(one);
                            break;
                        }
                    }
                    continue;
                }

                var xtotal_temp = xInlList.Sum(x => x.Money);

                //获取最后的级别， 按金额大到小逐一剔除记录，直到最接近总额度
                foreach (var one in xExpList)
                {
                    //剔除某条记录
                    xExpList = xExpList.Where(x => x != one).OrderByDescending(r => r.Money).ToList();

                    //判断剩下总金额是否超过额度
                    if (xLimitTotal < xtotal_temp + xExpList.Sum(x => x.Money))
                    {
                        continue;
                    }
                    else
                    {
                        xInlList.AddRange(xExpList);
                        break;
                    }

                }

                //新案总额
                Total = xInlList.Sum(x => x.Money);

                InlList.AddRange(xInlList);
            }


            //如果选中按档排除月
            if (checkBox3.CheckState == CheckState.Checked)
            {
                bankdatalist = bankdatalist.Where(x => x.InTime.Year == dateTimePicker1.Value.Year).OrderByDescending(r => r.Money).ToList();
            }


            //90后 +4501
            var Level1 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("199") && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

            //90后 +非4501
            var Level2 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("199") && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();


            //85后+4501
            var Level3 = bankdatalist.Where(x => (x.ID.Substring(6, 4).Equals("1985") || x.ID.Substring(6, 4).Equals("1986") || x.ID.Substring(6, 4).Equals("1987") || x.ID.Substring(6, 4).Equals("1988") || x.ID.Substring(6, 4).Equals("1989"))

                                                && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

            //85后 +非4501
            var Level4 = bankdatalist.Where(x => (x.ID.Substring(6, 4).Equals("1985") || x.ID.Substring(6, 4).Equals("1986") || x.ID.Substring(6, 4).Equals("1987") || x.ID.Substring(6, 4).Equals("1988") || x.ID.Substring(6, 4).Equals("1989"))

                                            && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();

            //80-84 +4501
            var Level5 = bankdatalist.Where(x => (x.ID.Substring(6, 4).Equals("1980") || x.ID.Substring(6, 4).Equals("1981") || x.ID.Substring(6, 4).Equals("1982") || x.ID.Substring(6, 4).Equals("1983") || x.ID.Substring(6, 4).Equals("1984"))

                                                && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

            //80-84  +非4501
            var Level6 = bankdatalist.Where(x => (x.ID.Substring(6, 4).Equals("1980") || x.ID.Substring(6, 4).Equals("1981") || x.ID.Substring(6, 4).Equals("1982") || x.ID.Substring(6, 4).Equals("1983") || x.ID.Substring(6, 4).Equals("1984"))

                                            && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();

            //70后 +4501
            var Level7 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("197") && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

            //70后 +非4501
            var Level8 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("197") && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();


            //60后 +4501
            var Level9 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("196") && x.ID.StartsWith("4501")).OrderByDescending(r => r.Money).ToList();

            //60后 +非4501
            var Level10 = bankdatalist.Where(x => x.ID.Substring(6, 3).Equals("196") && x.ID.Substring(0, 4) != "4501").OrderByDescending(r => r.Money).ToList();


            for (int i = 0; i < 10; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    switch (i)
                    {
                        case 0: AllLevel[i] = Level1; break;
                        case 1: AllLevel[i] = Level2; break;
                        case 2: AllLevel[i] = Level3; break;
                        case 3: AllLevel[i] = Level4; break;
                        case 4: AllLevel[i] = Level5; break;
                        case 5: AllLevel[i] = Level6; break;
                        case 6: AllLevel[i] = Level7; break;
                        case 7: AllLevel[i] = Level8; break;
                        case 8: AllLevel[i] = Level9; break;
                        case 9: AllLevel[i] = Level10; break;
                        default: break;
                    }
                }
                else
                {

                }
            }



            //日期优先的数据超额度，则不进行按年龄筛选的机制，直接进入逐项剔除机制
            //if (Level0.Sum(x => x.Money) > LimitTotal)
            //{
            //    ContinueFlag = false;
            //    ExpList.AddRange(Level0);
            //}
            //else
            //{
            //    InlList.AddRange(Level0);
            //    Total = Level0.Sum(x => x.Money);
            //}


            // 优先获取指定日期的数据后，仍不超额度的情况下，继续获取其他数据
            if (ContinueFlag)
            {
                //按优先级高到低(Level1~Level10)取数
                foreach (var one in AllLevel)
                {
                    if (one != null && one.Count != 0)
                    {
                        Total = Total + one.Sum(x => x.Money);

                        if (Total < LimitTotal)
                        {
                            InlList.AddRange(one);
                            continue;
                        }
                        else
                        {
                            ExpList.AddRange(one);
                            break;
                        }
                    }
                    continue;
                }
            }

            var total_temp = InlList.Sum(x => x.Money);

            //获取最后的级别， 按金额大到小逐一剔除记录，直到最接近总额度
            foreach (var one in ExpList)
            {
                //剔除某条记录
                ExpList = ExpList.Where(x => x != one).OrderByDescending(r => r.Money).ToList();

                //判断剩下总金额是否超过额度
                if (LimitTotal < total_temp + ExpList.Sum(x => x.Money))
                {
                    continue;
                }
                else
                {
                    InlList.AddRange(ExpList);
                    break;
                }

            }

            //InlList = InlList.AddRange(xInlList);


            //MessageBox.Show(InlList.Sum(x => x.Money).ToString());


            //var Level1Amt = Level1.Sum(x => x.Money);
            //var Level2Amt = Level2.Sum(x => x.Money);
            //var Level3Amt = Level3.Sum(x => x.Money);
            //var Level4Amt = Level4.Sum(x => x.Money);
            //var Level5Amt = Level5.Sum(x => x.Money);
            //var Level6Amt = Level6.Sum(x => x.Money);
            //var Level7Amt = Level7.Sum(x => x.Money);
            //var Level8Amt = Level8.Sum(x => x.Money);
            //var Level9Amt = Level9.Sum(x => x.Money);
            //var Level10Amt = Level10.Sum(x => x.Money);

            //var Total = Level1Amt + Level2Amt + Level3Amt + Level4Amt  + Level5Amt + Level6Amt + Level7Amt + Level8Amt + Level9Amt + Level10Amt;

            //MessageBox.Show(Total.ToString());

            //导出Excel
            //DataSetToExcel(data, textBox2.Text);

            DataSetToExcel(InlList.OrderByDescending(r => r.Money).ToList(), textBox2.Text);

            //打开Excel
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;

            Workbook book = app.Workbooks.Open(textBox2.Text);

            MessageBox.Show("抢案结果处理完成！");


            if (string.IsNullOrEmpty(textBox6.Text))
            {
                MessageBox.Show("请选择完整结果模板");
                return;
            }

            DataSetToExcel(InlList, 2, 1, 1, textBox6.Text);
            DataSetToExcel(ExceptList, 2, 1, 2, textBox6.Text);
            DataSetToExcel(ExpList, 2, 1, 3, textBox6.Text);

            DataSetToExcel(Level0, 2, 1, 4, textBox6.Text);
            DataSetToExcel(Level1, 2, 1, 5, textBox6.Text);
            DataSetToExcel(Level2, 2, 1, 6, textBox6.Text);
            DataSetToExcel(Level3, 2, 1, 7, textBox6.Text);
            DataSetToExcel(Level4, 2, 1, 8, textBox6.Text);
            DataSetToExcel(Level5, 2, 1, 9, textBox6.Text);
            DataSetToExcel(Level6, 2, 1, 10, textBox6.Text);
            DataSetToExcel(Level7, 2, 1, 11, textBox6.Text);
            DataSetToExcel(Level8, 2, 1, 12, textBox6.Text);
            DataSetToExcel(Level9, 2, 1, 13, textBox6.Text);
            DataSetToExcel(Level10, 2, 1, 14, textBox6.Text);

            MessageBox.Show("完整结果处理完成！");
        }

        /// <summary>
        /// 读取Excel为DataSet
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="FilterStr"></param>
        /// <returns></returns>
        public DataSet ExcelToDS(string Path, string FilterStr)
        {
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            //OleDbConnection conn = new OleDbConnection(strConn);

            //获取文件扩展名
            string strExtension = System.IO.Path.GetExtension(Path);
            //Excel的连接
            OleDbConnection objConn = null;
            switch (strExtension)
            {
                case ".xls":
                    objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";" + "Extended Properties=\"Excel 8.0;\"");
                    break;
                case ".xlsx":
                    objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";" + "Extended Properties=\"Excel 12.0;\"");
                    break;
                default:
                    objConn = null;
                    break;
            }
            if (objConn == null)
            {
                return null;
            }

            objConn.Open();

            System.Data.DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            string tableName = schemaTable.Rows[0][2].ToString().Trim();

            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = FilterStr;
            myCommand = new OleDbDataAdapter(strExcel, objConn);
            ds = new DataSet();
            myCommand.Fill(ds, tableName);

            return ds;
        }


        /// <summary>
        /// 读取Excel为DataSet
        /// </summary>
        /// <param name="strExcelPath"></param>
        /// <param name="tableName"></param>
        /// <param name="ExcelFilterStr"></param>
        /// <returns></returns>
        public DataSet GetExcelTableByOleDB(string strExcelPath, string tableName, string ExcelFilterStr)
        {
            try
            {
                //数据表
                DataSet ds = new DataSet();
                //获取文件扩展名
                string strExtension = System.IO.Path.GetExtension(strExcelPath);
                //Excel的连接
                OleDbConnection objConn = null;
                switch (strExtension)
                {
                    case ".xls":
                        objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;\"");
                        break;
                    case ".xlsx":
                        objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 12.0;\"");
                        break;
                    default:
                        objConn = null;
                        break;
                }
                if (objConn == null)
                {
                    return null;
                }
                objConn.Open();

                OleDbDataAdapter myCommand = null;
                myCommand = new OleDbDataAdapter(ExcelFilterStr, objConn);
                myCommand.Fill(ds, tableName);

                objConn.Close();

                return ds;
            }
            catch
            {
                return null;
            }
        }


        /// <summary>
        /// 通过关键字获取1个文件的路径
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public string GetOneFileByKeyWord(string Path, string FilesType, string keyword)
        {
            DirectoryInfo dir = new DirectoryInfo(Path);
            //path为某个目录，如： “D:\Program Files”
            FileInfo[] inf = dir.GetFiles();
            foreach (FileInfo finf in inf)
            {
                if (finf.Extension.Equals(FilesType) && finf.Name.Contains(keyword))
                    return (Path + "\\" + finf.Name);
            }

            return string.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str = GetOneFileByKeyWord(textBox4.Text, ".xls", textBox3.Text);

            if (!string.IsNullOrEmpty(str))
            {
                textBox1.Text = str;

                //彻底删除加密文件
                string gpgPath = GetOneFileByKeyWord(textBox4.Text, ".gpg", "");

                if (!string.IsNullOrEmpty(gpgPath))
                {
                    DeleteFile(gpgPath);
                }

                //彻底删除压缩包文件
                string ZipPath = GetOneFileByKeyWord(textBox4.Text, ".zip", "");

                if (!string.IsNullOrEmpty(ZipPath))
                {
                    DeleteFile(ZipPath);
                }
            }
        }

        //将数据写入已存在Excel
        public static void writeExcel(string result, string filepath)
        {
            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(filepath,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];

            //在第一列的左边插入一列  1:第一列
            //xlShiftToRight:向右移动单元格   xlShiftDown:向下移动单元格
            //Range Columns = (Range)xSheet.Columns[1, System.Type.Missing];
            //Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);

            //4.向相应对位置写入相应的数据
            xSheet.Cells[1][1] = result;

            //5.保存保存WorkBook
            xBook.Save();
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;
        }

        /// <summary>
        /// 保存DataSet到Excel
        /// </summary>
        /// <param name="dataSet"></param>
        /// <param name="Path"></param>
        public static void DataSetToExcel(DataSet dataSet, string Path)
        {
            System.Data.DataTable dataTable = dataSet.Tables[0];
            int rowNumber = dataTable.Rows.Count;

            int rowIndex = 2;
            int colIndex = 0;


            if (rowNumber == 0)
            {
                return;
            }

            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(Path,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];

            //var Allco = dataTable.Columns;

            ////生成字段名称
            //foreach (DataColumn col in Allco)
            //{
            //    colIndex++;
            //    xSheet.Cells[1,colIndex] = col.ColumnName;

            //}

            //填充数据 
            foreach (DataRow row in dataTable.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in dataTable.Columns)
                {
                    colIndex++;
                    xSheet.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                }
                xSheet.Cells[rowIndex, colIndex + 1] = "恒远";
                xSheet.Cells[rowIndex, colIndex + 2] = "EI5";
            }

            //设置边框区域
            Range titleRange = xSheet.Range[xSheet.Cells[3, 1], xSheet.Cells[rowIndex, 4]];
            titleRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置边框  
            titleRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;//边框常规粗细  

            //5.保存保存WorkBook
            xBook.Save();
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;

        }

        public class KeyMyExcelProcess
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                try
                {
                    IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
                    int k = 0;
                    GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                    p.Kill();     //关闭进程k
                }
                catch (System.Exception ex)
                {
                    throw ex;
                }
            }
        }


        public static void DataSetToExcel(List<BankDataObj> result, string Path)
        {
            //System.Data.DataTable dataTable = dataSet.Tables[0];
            //int rowNumber = dataTable.Rows.Count;

            //int rowIndex = 2;
            //int colIndex = 0;


            //if (rowNumber == 0)
            //{
            //    return;
            //}

            if (result.Count == 0)
            {
                return;
            }

            int rowIndex = 3;

            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(Path,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];


            //获取编辑范围
            Range range;
            for (int i = 0; i < 1000; i++)
            {
                //range = (Microsoft.Office.Interop.Excel.Range)xSheet.Rows[rowIndex + 1+i, Missing.Value];
                range = (Microsoft.Office.Interop.Excel.Range)xSheet.Rows[3 + i, Type.Missing];
                range.Delete();
                //range.EntireRow.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
            }


            //var Allco = dataTable.Columns;

            ////生成字段名称
            //foreach (DataColumn col in Allco)
            //{
            //    colIndex++;
            //    xSheet.Cells[1,colIndex] = col.ColumnName;

            //}

            //填充数据 
            //foreach (DataRow row in dataTable.Rows)
            //{
            //    rowIndex++;
            //    colIndex = 0;
            //    foreach (DataColumn col in dataTable.Columns)
            //    {
            //        colIndex++;
            //        xSheet.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
            //    }
            //    xSheet.Cells[rowIndex, colIndex + 1] = "恒远";
            //    xSheet.Cells[rowIndex, colIndex + 2] = "EI5";
            //}


            int k = 0;
            var value = (string)xSheet.Cells[rowIndex, 1];

            //查找起始填充数据的单元格行号
            while (!string.IsNullOrEmpty(value))
            {
                k++;
                value = (string)xSheet.Cells[rowIndex + k, 1];
            }

            rowIndex = rowIndex + k;

            foreach (var one in result)
            {
                xSheet.Cells[rowIndex, 1] = one.CustomerCode;
                xSheet.Cells[rowIndex, 2] = one.Money;
                xSheet.Cells[rowIndex, 3] = "恒远";
                xSheet.Cells[rowIndex, 4] = "EI5";
                rowIndex++;
            }

            //设置边框区域
            Range titleRange = xSheet.Range[xSheet.Cells[3, 1], xSheet.Cells[rowIndex, 4]];
            titleRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置边框  
            titleRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;//边框常规粗细  

            //5.保存保存WorkBook
            xBook.Save();
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;

        }


        /// <summary>
        /// 案列数，偏移量，目标Sheet,固定字符串，路径生成结果
        /// </summary>
        /// <param name="result"></param>
        /// <param name="colIndex"></param>
        /// <param name="OffsetX"></param>
        /// <param name="OffsetY"></param>
        /// <param name="Path"></param>
        public static void DataSetToExcel(List<BankDataObj> result, int OffsetRow, int OffsetCol, int SheetNO, string Path)
        {
            if (result.Count == 0)
            {
                return;
            }
            int rowIndex = OffsetRow;

            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(Path,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[SheetNO];

            foreach (var one in result)
            {
                rowIndex++;
                xSheet.Cells[rowIndex, OffsetCol] = one.Name;
                xSheet.Cells[rowIndex, OffsetCol + 1] = one.Money;
                xSheet.Cells[rowIndex, OffsetCol + 2] = one.CustomerCode;
                xSheet.Cells[rowIndex, OffsetCol + 3] = one.ID;
                xSheet.Cells[rowIndex, OffsetCol + 4] = one.City;
                xSheet.Cells[rowIndex, OffsetCol + 5] = one.Province;
                xSheet.Cells[rowIndex, OffsetCol + 6] = one.InTime;
            }

            xSheet.Cells[rowIndex + 1, OffsetCol] = "总金额:";
            xSheet.Cells[rowIndex + 1, OffsetCol + 1] = result.Sum(s => s.Money);

            //设置边框区域
            Range titleRange = xSheet.Range[xSheet.Cells[OffsetRow, OffsetCol], xSheet.Cells[rowIndex, OffsetCol + 6]];
            titleRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置边框  
            titleRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;//边框常规粗细  

            //5.保存保存WorkBook
            xBook.Save();
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;

        }


        /// <summary>
        /// 根据路径删除文件(彻底删除)
        /// </summary>
        /// <param name="path"></param>
        public void DeleteFile(string path)
        {
            FileAttributes attr = File.GetAttributes(path);
            if (attr == FileAttributes.Directory)
            {
                Directory.Delete(path, true);
            }
            else
            {
                File.Delete(path);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();     //显示选择文件对话框
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);   //默认打开当前用户桌面路径
            openFileDialog1.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";  //只选取Excel文件
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog1.FileName;          //显示文件路径
            }
        }


        /// <summary>
        /// Decompression the zip file 
        /// </summary>
        /// <param name="zipFilePath">the source path of the zip file</param>
        /// <param name="savePath">What path you want to save the directory that was decompressioned</param>
        private static void UnZipFile(string zipFilePath, string savePath)
        {
            if (!File.Exists(zipFilePath))
            {
                //Console.WriteLine("Cannot find file '{0}'", zipFilePath);
                MessageBox.Show("错误！！！ 压缩包文件不存在，不能解压！");
                return;
            }

            using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipFilePath)))
            {

                ZipEntry theEntry;
                while ((theEntry = s.GetNextEntry()) != null)
                {

                    string fullPath = JudgeFullPath(savePath) + theEntry.Name;
                    Console.WriteLine(fullPath);

                    string directoryName = Path.GetDirectoryName(fullPath);
                    string fileName = Path.GetFileName(fullPath);

                    //if (String.IsNullOrEmpty(savePath))
                    //{
                    //    directoryName = savePath + "/"+directoryName;
                    //}
                    // create directory
                    if (directoryName.Length > 0)
                    {
                        Directory.CreateDirectory(directoryName);
                    }

                    if (fileName != String.Empty)
                    {
                        using (FileStream streamWriter = File.Create(fullPath))
                        {

                            int size = 2048;
                            byte[] data = new byte[2048];
                            while (true)
                            {
                                size = s.Read(data, 0, data.Length);
                                if (size > 0)
                                {
                                    streamWriter.Write(data, 0, size);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Judge the last symbol of the full path wether is "/" or not.
        /// </summary>
        /// <param name="_path"></param>
        /// <returns></returns>
        private static string JudgeFullPath(string _path)
        {
            if (!string.IsNullOrEmpty(_path) && _path.Length > 4)
            {
                string lastSymbol = _path.Substring(_path.Length - 1);
                if (lastSymbol == "\\")
                {
                    return _path;
                }
                return _path + "\\";
            }
            return _path;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //string zipFilePath = GetOneFileByKeyWord(textBox4.Text,".zip",textBox3.Text);
            //UnZipFile(zipFilePath, textBox4.Text);
            if (pTimer.Enabled)
            {
                button6.Text = "自动生成结果(已停止)";
                pTimer.Enabled = false;
            }
            else
            {
                button6.Text = "自动生成结果(已启动)";
                pTimer.Enabled = true;
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.CheckState == CheckState.Checked)
            {
                checkBox2.CheckState = CheckState.Unchecked;
                checkBox3.CheckState = CheckState.Unchecked;
            }

            if (checkBox1.CheckState == CheckState.Unchecked && checkBox2.CheckState == CheckState.Unchecked && checkBox3.CheckState == CheckState.Unchecked)
            {
                dateTimePicker1.Enabled = false;
            }

            if (checkBox1.CheckState == CheckState.Checked || checkBox2.CheckState == CheckState.Checked || checkBox3.CheckState == CheckState.Checked)
            {
                dateTimePicker1.Enabled = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog4 = new OpenFileDialog();     //显示选择文件对话框
            openFileDialog4.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);   //默认打开当前用户桌面路径
            openFileDialog4.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";  //只选取Excel文件
            openFileDialog4.FilterIndex = 2;
            openFileDialog4.RestoreDirectory = true;

            if (openFileDialog4.ShowDialog() == DialogResult.OK)
            {
                textBox6.Text = openFileDialog4.FileName;          //显示文件路径
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox2.CheckState == CheckState.Unchecked && checkBox3.CheckState == CheckState.Unchecked)
            {
                dateTimePicker1.Enabled = false;
            }

            if (checkBox2.CheckState == CheckState.Checked || checkBox3.CheckState == CheckState.Checked)
            {
                dateTimePicker1.Enabled = true;
            }

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.CheckState == CheckState.Unchecked && checkBox3.CheckState == CheckState.Unchecked)
            {
                dateTimePicker1.Enabled = false;
            }

            if (checkBox2.CheckState == CheckState.Checked || checkBox3.CheckState == CheckState.Checked)
            {
                dateTimePicker1.Enabled = true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog5 = new OpenFileDialog();     //显示选择文件对话框
            openFileDialog5.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);   //默认打开当前用户桌面路径
            openFileDialog5.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";  //只选取Excel文件
            openFileDialog5.FilterIndex = 2;
            openFileDialog5.RestoreDirectory = true;

            if (openFileDialog5.ShowDialog() == DialogResult.OK)
            {
                this.textBox10.Text = openFileDialog5.FileName;          //显示文件路径
            }
        }
    }
}