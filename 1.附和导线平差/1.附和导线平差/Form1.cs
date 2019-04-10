using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;//输入输出
using Excel = Microsoft.Office.Interop.Excel;//Excel表格
using System.Drawing.Drawing2D;//绘图

namespace _1.附和导线平差
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #region 时间控件
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView3.AllowUserToAddRows = false;

            toolStripStatusLabel3.Text = DateTime.Now.ToString();
            timer1.Enabled = true;
            timer1.Interval = 1000;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel3.Text = DateTime.Now.ToString();
        }
        #endregion
        #region 定义主要变量
        int k;//左角右角
        List<string> dianhao;//点号
        List<double> Xzuobiao;//X坐标
        List<double> Yzuobiao;//Y坐标
        List<double> guancejiao;//观测角
        List<double> juli;//距离
        List<double> fangweijiao;//方位角
        List<double> jiaogaizheng;//角度改正数
        List<double> jiaogaihou;//改正后角值
        List<double> Xzengliang;//X坐标增量
        List<double> Yzengliang;//Y坐标增量
        List<double> Xgaizheng;//X坐标改正值
        List<double> Ygaizheng;//Y坐标改正值
        List<double> Xgaihou;//改正后X坐标增量
        List<double> Ygaihou;//改正后Y坐标增量
        double jiaoduBHC;//角度闭合差
        double XBHC;//X坐标增量闭合差
        double YBHC;//Y坐标增量闭合差
        Bitmap image;
        #endregion
        #region 初始化
        public void chushihua()
        {
            dianhao = new List<string>();
            Xzuobiao = new List<double>();
            Yzuobiao = new List<double>();
            guancejiao = new List<double>();
            juli = new List<double>();
            fangweijiao = new List<double>();
            jiaogaizheng = new List<double>();
            jiaogaihou = new List<double>();
            Xzengliang = new List<double>();
            Yzengliang = new List<double>();
            Xgaizheng = new List<double>();
            Ygaizheng = new List<double>();
            Xgaihou = new List<double>();
            Ygaihou = new List<double>();
        }
        #endregion
        #region 文件打开
        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();//清空表一的元素
            dataGridView2.Rows.Clear();
            txt_qishi.Text = "";//清空方位角
            txt_zongzhi.Text = "";

            openFileDialog1.Title = "附和导线数据打开";
            openFileDialog1.Filter = "文本文件(*.txt)|*.txt|Excel旧版本文件(.xls)|*.xls|Excel新版本文件(*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                #region txt文件
                if (openFileDialog1.FilterIndex == 1)
                {
                    StreamReader sr = new StreamReader(openFileDialog1.FileName,Encoding.Default);
                    sr.ReadLine();//跳过第一行
                    string[] str = sr.ReadLine().Split(',');//读取第二行
                    int i = 0, j = 0;
                    while (!sr.EndOfStream)
                    {
                        if (str[0] != "观测数据：")
                        {
                            dataGridView1.Rows.Add();
                            for (int q = 0; q < str.Length; q++)
                            {
                                dataGridView1.Rows[i].Cells[q].Value = str[q];
                            }
                            str = sr.ReadLine().Split(',');
                            i++;
                        }
                        else
                        {
                            while (!sr.EndOfStream)
                            {
                                str = sr.ReadLine().Split(',');
                                dataGridView2.Rows.Add();
                                for (int q = 0; q < str.Length; q++)
                                {
                                    dataGridView2.Rows[j].Cells[q].Value = str[q];
                                }
                                j++;
                            }
                            break;//跳出当前循环
                        }
                    }
                    sr.Close();//退出文件流
                }
                #endregion
                #region excel文件
                else 
                {
                    Excel.Application excel = new Excel.Application();
                    excel.Visible = false;//以只读方式打开Excel文件
                    Excel.Workbook wk = excel.Workbooks.Open(openFileDialog1.FileName);
                    Excel.Worksheet ws = excel.Workbooks[1].Worksheets[1];
                    int rows = ws.UsedRange.Rows.Count;//获取非空的列数
                    int columns = ws.UsedRange.Columns.Count;//获取非空的行数
                    for (int i = 0; i < 4; i++)//已知数据必须占用4个行，否则要出错，已知方位角则可以让A,D为空
                    {
                        dataGridView1.Rows.Add();
                        for (int j = 0; j < columns; j++ )
                        {
                            dataGridView1.Rows[i].Cells[j].Value = ws.Cells[i + 2, j + 1].Value;
                        }
                    }
                    for (int i = 0; i < rows - 6; i++)
                    {
                        dataGridView2.Rows.Add();
                        for (int j = 0; j < columns; j++)
                        { 
                            dataGridView2.Rows[i].Cells[j].Value = ws.Cells[i + 7, j + 1].Value;//单元格为空也能赋值
                        }
                    }
                    wk.Close();//退出文件流
                }
                #endregion
            }
        }
        #endregion
        #region 导线计算
        private void 计算ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //计算过程中角度全部以弧度的形式参与计算
            chushihua();//变量初始化
            dataGridView3.Rows.Clear();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView2.AllowUserToAddRows = false;
            #region 判断左角右角
            if (rdb_left.Checked)
            {
                k = 1;
            }
            else
            {
                k = -1;
            }
            #endregion
            #region 数据导入
            //文件里面不能为空，一旦为空，无法计算
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (i == 0 || i == dataGridView1.Rows.Count - 1)
                    {
                        dianhao.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                    }
                    Xzuobiao.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", "")));//replace去掉数据中的空格
                    Yzuobiao.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", "")));
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    dianhao.Insert(i + 1,dataGridView2.Rows[i].Cells[0].Value.ToString());
                    guancejiao.Add(Caculates.dmstohudu(Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value.ToString().Replace(" ", ""))));//弧度
                    if (dataGridView2.Rows[i].Cells[2].Value != null)
                    {
                        juli.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value.ToString().Replace(" ", "")));
                    }
                }
            }
            catch
            {
                MessageBox.Show("请输入正确的数据！");
                return;
            }
            #endregion
            #region 计算方位角
            if (txt_qishi.Text == "" || txt_zongzhi.Text == "")
            {
                fangweijiao.Add(Caculates.fangwei(Xzuobiao[0], Yzuobiao[0], Xzuobiao[1], Yzuobiao[1]));//返回弧度值
                fangweijiao.Add(Caculates.fangwei(Xzuobiao[2], Yzuobiao[2], Xzuobiao[3], Yzuobiao[3]));
                txt_qishi.Text = Caculates.hudutodms(fangweijiao[0]).ToString();
                txt_zongzhi.Text = Caculates.hudutodms(fangweijiao[1]).ToString();
            }
            else
            {
                fangweijiao.Add(Caculates.dmstohudu(Convert.ToDouble(txt_qishi.Text)));
                fangweijiao.Add(Caculates.dmstohudu(Convert.ToDouble(txt_zongzhi.Text)));
            }
            #endregion
            #region 已知数据导出
            for (int i = 0; i < dianhao.Count; i++)//点号
            {
                dataGridView3.Rows.Add();
                dataGridView3.Rows[i].Cells[0].Value = dianhao[i];
            }
            for (int i = 0; i < guancejiao.Count; i++)//观测角
            {
                dataGridView3.Rows[i + 1].Cells[1].Value = Caculates.hudutodms(guancejiao[i]);
            }
            for (int i = 0; i < juli.Count; i++)//距离
            {
                dataGridView3.Rows[i + 1].Cells[5].Value = juli[i];
            }
            #endregion
            #region 角度闭合差
            //提前判断加减180，因为加减180无法直接用公式判断
            List<double> fangwei1 = new List<double>();
            fangwei1.Add(fangweijiao[0]);
            double n = 0;//计算大于360或者小于0的值的累积和
            for (int i = 0; i < guancejiao.Count; i++)
            {
                double a = fangwei1[i] + k * guancejiao[i] - k * Math.PI;
                if (a > Math.PI * 2)
                {
                    a = a - Math.PI * 2;
                    n = n - Math.PI * 2;
                }
                else if (a < 0)
                {
                    a = a + Math.PI * 2;
                    n = n + Math.PI * 2;
                }
                fangwei1.Add(a);
            }

            jiaoduBHC = fangweijiao[0] + k * guancejiao.Sum() - fangweijiao[1] - k * Math.PI * (guancejiao.Count) + n;//观测值减去真实值
            if (Caculates.hudutos(jiaoduBHC) > 40 * Math.Sqrt(guancejiao.Count))//限差设置为40倍的根号n
            {
                MessageBox.Show("角度闭合差超限！！！");
            }
            #endregion
            #region 方位角计算
            for (int i = 0; i < guancejiao.Count; i++)
            {
                jiaogaizheng.Add(- k * jiaoduBHC / guancejiao.Count);//弧度
                jiaogaihou.Add(guancejiao[i] + jiaogaizheng[i]);
                double a = fangweijiao[i] + k * jiaogaihou[i] - k * Math.PI;
                if (a > Math.PI * 2)
                {
                    a = a - Math.PI * 2;
                }
                else if (a < 0)
                {
                    a = a + Math.PI * 2;
                }
                fangweijiao.Insert(i + 1, a);
            }

            for (int i = 0; i < guancejiao.Count; i++)
            {
                dataGridView3.Rows[i + 1].Cells[2].Value = Caculates.hudutos(jiaogaizheng[i]);//改正数
                dataGridView3.Rows[i + 1].Cells[3].Value = Caculates.hudutodms(jiaogaihou[i]);//改正后角值
            }
            dataGridView3.Rows[guancejiao.Count + 1].Cells[2].Value = Caculates.hudutos(jiaoduBHC);//角度闭合差
            fangweijiao.RemoveAt(guancejiao.Count + 1);//方位角多出一个，所以删除，不用判断，计算机不会算错
            for (int i = 0; i < fangweijiao.Count; i++)
            {
                dataGridView3.Rows[i].Cells[4].Value = Caculates.hudutodms(fangweijiao[i]);//方位角
            }
            #endregion
            #region 计算坐标增量
            for (int i = 0; i < juli.Count; i++)
            {
                Xzengliang.Add(juli[i] * Math.Cos(fangweijiao[i + 1]));
                Yzengliang.Add(juli[i] * Math.Sin(fangweijiao[i + 1]));
            }
            XBHC = Xzengliang.Sum() - (Xzuobiao[2] - Xzuobiao[1]);//观测值减去真实值
            YBHC = Yzengliang.Sum() - (Yzuobiao[2] - Yzuobiao[1]);
            double aa = Math.Sqrt((XBHC * XBHC + YBHC * YBHC)) / juli.Sum();
            if (aa > 0.00025)//限差设置为1/4000
            {
                MessageBox.Show("导线全长闭合差超限！");
            }
            for (int i = 0; i < juli.Count; i++)
            {
                Xgaizheng.Add(XBHC * juli[i] / juli.Sum());
                Ygaizheng.Add(YBHC * juli[i] / juli.Sum());
                Xgaihou.Add(Xzengliang[i] - Xgaizheng[i]);
                Ygaihou.Add(Yzengliang[i] - Ygaizheng[i]);
            }

            for (int i = 0; i < Xzengliang.Count; i++)
            {
                dataGridView3.Rows[i + 1].Cells[6].Value = Math.Round(Xzengliang[i], 4);//坐标增量
                dataGridView3.Rows[i + 1].Cells[7].Value = Math.Round(Yzengliang[i], 4);
                dataGridView3.Rows[i + 1].Cells[8].Value = Math.Round(Xgaizheng[i], 4) * 100;//坐标增量改正数
                dataGridView3.Rows[i + 1].Cells[9].Value = Math.Round(Ygaizheng[i], 4) * 100;
                dataGridView3.Rows[i + 1].Cells[10].Value = Math.Round(Xgaihou[i], 4);//改后坐标增量
                dataGridView3.Rows[i + 1].Cells[11].Value = Math.Round(Ygaihou[i], 4);
            }
            dataGridView3.Rows[Xzengliang.Count + 2].Cells[8].Value = Math.Round(XBHC, 4) * 100;//坐标增量闭合差
            dataGridView3.Rows[Xzengliang.Count + 2].Cells[9].Value = Math.Round(YBHC, 4) * 100;
            #endregion
            #region 计算坐标
            for (int i = 0; i < Xgaihou.Count - 1; i++)//坐标计算c点多出一个，所以-1，不用判断，计算机不会算错
            {
                Xzuobiao.Insert(i + 2, Xgaihou[i] + Xzuobiao[i + 1]);
                Yzuobiao.Insert(i + 2, Ygaihou[i] + Yzuobiao[i + 1]);
            }
            for (int i = 0; i < Xzuobiao.Count; i++)
            {
                dataGridView3.Rows[i].Cells[12].Value = Math.Round(Xzuobiao[i], 4);
                dataGridView3.Rows[i].Cells[13].Value = Math.Round(Yzuobiao[i], 4);
            }
            #endregion
        }
        #endregion
        #region 文件保存
        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "附和导线平差计算结果保存";
            saveFileDialog1.Filter = "文本文件(*.txt)|*.txt|Excel旧版本文件(*.xls)|*.xls|Excel新版本文件(*.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                #region txt文件
                if (saveFileDialog1.FilterIndex == 1)
                {
                    StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
                    List<string> arrstr = new List<string>();
                    string str = null;
                    sw.WriteLine("附和导线近似平差计算结果：\n");
                    for (int i = 0; i < dataGridView3.Columns.Count; i++)//表头
                    {
                        arrstr.Add(string.Format("{0,-8}",dataGridView3.Columns[i].HeaderText));//-8表示格式化8个字符，原字符左对齐，不足则补空格,但是对文字好像不是特别好使
                    }
                    str = string.Join("\t", arrstr);
                    sw.WriteLine(str);
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)//数据
                    {
                        str = null;
                        arrstr.Clear();
                        for (int j = 0; j < dataGridView3.Columns.Count; j++)
                        {
                            arrstr.Add(string.Format("{0,-8}",dataGridView3.Rows[i].Cells[j].Value));
                        }
                        str = string.Join("\t", arrstr);
                        sw.WriteLine(str);
                    }
                    sw.Close();
                }
                #endregion
                #region excel文件
                else
                {
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wk = excel.Workbooks.Add(true);//为excel对象添加一个工作簿
                    Excel.Worksheet ws = excel.Workbooks[1].Worksheets[1];
                    for (int i = 0; i < dataGridView3.Columns.Count; i++)//表头
                    {
                        ws.Cells[1, i + 1].Value = dataGridView3.Columns[i].HeaderText;
                    }
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)//数据
                    {
                        for (int j = 0; j < dataGridView3.Columns.Count; j++)
                        { 
                            ws.Cells[i + 2, j + 1].Value = dataGridView3.Rows[i].Cells[j].Value;
                        }
                    }
                    ws.Columns.AutoFit();//自动调整列宽
                    ws.SaveAs(saveFileDialog1.FileName);//保存工作表
                    wk.Close();
                }
                #endregion
                MessageBox.Show("保存成功！");
            }
        }
        #endregion
        #region 绘制图形
        private void 绘图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Pen p = new Pen(Color.Black, 2.5f);
            Pen p1 = new Pen(Color.Red, 3);
            Pen p2 = new Pen(Color.Blue, 2);
            image = new Bitmap((int)(Yzuobiao.Max() - Yzuobiao.Min()) + 200, (int)(Xzuobiao.Max() - Xzuobiao.Min()) + 200);//显示图形范围
            Graphics g = Graphics.FromImage(image);
            g.RotateTransform(-90);//旋转为测量坐标系
            g.TranslateTransform(-(int)(Xzuobiao.Max() + 100), -(int)Yzuobiao.Min() + 100);//划定原点位置
            PointF[] pf = new PointF[Xzuobiao.Count];
            //线形绘制
            for (int i = 0; i < Xzuobiao.Count; i++)
            {
                pf[i].X = (float)Xzuobiao[i];
                pf[i].Y = (float)Yzuobiao[i];
            }
            g.DrawLines(p, pf);

            //注记双线
            float[] single = { 0, 0.25f, 0.75f, 1 };
            p1.CompoundArray = single;
            g.DrawLine(p1, pf[0], pf[1]);
            g.DrawLine(p1, pf[pf.Length - 2], pf[pf.Length - 1]);

            //绘制三角
            Caculates.sanjiao(g, pf[0]);
            Caculates.sanjiao(g, pf[1]);
            Caculates.sanjiao(g, pf[pf.Length - 2]);
            Caculates.sanjiao(g, pf[pf.Length - 1]);

            //绘制圆弧
            for (int i = 0; i < guancejiao.Count; i++)
            {
                g.DrawArc(p2, pf[i + 1].X - 15, pf[i + 1].Y - 15, 30, 30, (float)(fangweijiao[i + 1] * 180 / Math.PI), -k * (float)(jiaogaihou[i] * 180 / Math.PI));//弧度必须化度
            }

            //绘制字体
            for (int i = 0; i < dianhao.Count; i++)
            {
                Caculates.ziti(g, pf[i], dianhao[i].ToString());
            }
            GraphicsState gstate = g.Save();
            g.ResetTransform();
            g.Restore(gstate);

            //绘制箭头
            AdjustableArrowCap cap = new AdjustableArrowCap(10, 20);
            p.CustomEndCap = cap;
            g.DrawLine(p, pf[pf.Length - 1].X + 50, pf[pf.Length - 1].Y + 50, pf[pf.Length - 1].X + 150, pf[pf.Length - 1].Y + 50);
            Caculates.ziti(g, new PointF(pf[pf.Length - 1].X + 150 + 25, pf[pf.Length - 1].Y + 50 + 25), "X");
            pictureBox1.Image = (Image)image;
        }
        #endregion
        #region bmp图形保存
        private void bmp图形ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "图像文件(*.bmp)|*.bmp";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                image.Save(saveFileDialog1.FileName);
            }
            MessageBox.Show("保存成功！");
        }
        #endregion
        #region dxf图形保存
        private void dxf图形ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "保存dxf文件";
            saveFileDialog1.Filter = "AutoCAD dxf文件(*.dxf)|*.dxf";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                {
                    sw.Write(Caculates.shiqian());
                    sw.Write("2\nENTITIES\n");
                    #region 绘制多线
                    sw.Write("0\nPOLYLINE\n");//多线段绘制，为一个整体线段
                    sw.Write("8\n");//图层
                    sw.Write("shiti\n");
                    sw.Write("66\n");//不太懂，应该是多线个数
                    sw.Write("1\n");
                    for (int i = 0; i < Xzuobiao.Count; i++)
                    {
                        sw.Write("0\nVERTEX\n");//多线段标识
                        sw.Write("8\n");//图层
                        sw.Write("shiti\n");
                        sw.Write("10\n");//X坐标
                        sw.Write(Yzuobiao[i] + "\n");
                        sw.Write("20\n");//Y坐标
                        sw.Write(Xzuobiao[i] + "\n");
                    }
                    sw.Write("0\nSEQEND\n");//多线段结束
                    #endregion
                    #region 双线
                    //用单线画出，不好看，可以考虑只画颜色不同的单线
                    for (int i = 0; i < Xzuobiao.Count - 1; i++)
                    {
                        if (i == 0 || i == Xzuobiao.Count - 2)
                        {
                            sw.Write("0\nLINE\n");//画线
                            sw.Write("8\n");//图层
                            sw.Write("qita\n");
                            sw.Write("10\n");//起始处X坐标
                            sw.Write(Yzuobiao[i] - 5 + "\n");
                            sw.Write("20\n");//起始处Y坐标
                            sw.Write(Xzuobiao[i] + "\n");
                            sw.Write("11\n");//终止处X坐标
                            sw.Write(Yzuobiao[i + 1] - 5 + "\n");
                            sw.Write("21\n");//终止处Y坐标
                            sw.Write(Xzuobiao[i + 1] + "\n");

                            sw.Write("0\nLINE\n");//画线
                            sw.Write("8\n");//图层
                            sw.Write("qita\n");
                            sw.Write("10\n");//起始处X坐标
                            sw.Write(Yzuobiao[i] + 5 + "\n");
                            sw.Write("20\n");//起始处Y坐标
                            sw.Write(Xzuobiao[i] + "\n");
                            sw.Write("11\n");//终止处X坐标
                            sw.Write(Yzuobiao[i + 1] + 5 + "\n");
                            sw.Write("21\n");//终止处Y坐标
                            sw.Write(Xzuobiao[i + 1] + "\n");
                        }
                    }
                    #endregion
                    #region 文字注记
                    for (int i = 0; i < Xzuobiao.Count; i++)
                    {
                        sw.Write(Caculates.zhuji(Yzuobiao[i], Xzuobiao[i], dianhao[i]));//注记
                    }
                    #endregion
                    #region 圆弧
                    for (int i = 0; i < guancejiao.Count; i++)
                    {
                        #region 右角
                        if (k == -1)
                        {
                            sw.Write("0\nARC\n");//单一圆
                            sw.Write("8\n");
                            sw.Write("zhuji\n");
                            sw.Write("10\n");//圆心X
                            sw.Write(Yzuobiao[i + 1] + "\n");
                            sw.Write("20\n");//圆心Y
                            sw.Write(Xzuobiao[i + 1] + "\n");
                            sw.Write("40\n");//圆的半径
                            sw.Write(15 + "\n");
                            sw.Write("50\n");//起始角，以度为单位，表示沿着X轴逆时针旋转的度数
                            sw.Write(630 - fangweijiao[i] * 180 / Math.PI + "\n");
                            sw.Write("51\n");//终止角，以度为单位，表示沿着X轴逆时针旋转的度数
                            sw.Write(630 - fangweijiao[i] * 180 / Math.PI + (jiaogaihou[i] * 180 / Math.PI) + "\n");
                        }
                        #endregion
                        #region 左角
                        else
                        {
                            sw.Write("0\nARC\n");//单一圆
                            sw.Write("8\n");
                            sw.Write("zhuji\n");
                            sw.Write("10\n");//圆心X
                            sw.Write(Yzuobiao[i + 1] + "\n");
                            sw.Write("20\n");//圆心Y
                            sw.Write(Xzuobiao[i + 1] + "\n");
                            sw.Write("40\n");//圆的半径
                            sw.Write(15 + "\n");
                            sw.Write("50\n");//起始角，以度为单位，表示沿着X轴逆时针旋转的度数
                            sw.Write(450 - fangweijiao[i + 1] * 180 / Math.PI + "\n");
                            sw.Write("51\n");//终止角，以度为单位，表示沿着X轴逆时针旋转的度数
                            sw.Write(450 - fangweijiao[i + 1] * 180 / Math.PI + (jiaogaihou[i] * 180 / Math.PI) + "\n");
                        }
                        #endregion
                    }
                    #endregion
                    sw.Write("0\nENDSEC\n");//第二段结束
                    sw.Write("0\nEOF\n");//文件结束
                }
                MessageBox.Show("保存成功！");
            }
        }
        #endregion
        #region 刷新
        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            txt_qishi.Text = "";
            txt_zongzhi.Text = "";
            dataGridView1.AllowUserToAddRows = true;
            dataGridView2.AllowUserToAddRows = true;
            pictureBox1.Image = null;//清空图形内容
        }
        #endregion
    }
}
