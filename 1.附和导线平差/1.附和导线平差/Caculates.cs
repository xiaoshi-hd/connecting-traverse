using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;//引用绘图
using System.Windows.Forms;//引用窗体控件

namespace _1.附和导线平差
{
    class Caculates
    {
        #region d.ms化弧度
        public static double dmstohudu(double dms)//角度在这里不会出现负数
        {
            double d, m, s;
            d = Math.Floor(dms);
            m = Math.Floor((dms - d) * 100);
            s = ((dms - d) * 100 - m) * 100;
            return (d + m / 60 + s / 3600) * Math.PI / 180;
        }
        #endregion
        //(输出才会用到化d.ms和s,所以不会影响计算的精度,只是影响输出表格中显示的精度)
        #region 弧度化d.ms
        public static double hudutodms(double hudu)//角度在这里不会出现负数
        {
            double d, m, s;
            double du = hudu * 180 / Math.PI;
            d = Math.Floor(du);
            m = Math.Floor((du - d) * 60);
            s = ((du - d) * 60 - m) * 60;
            return Math.Round(d + m / 100 + s / 10000, 4);//保留到秒
        }
        #endregion
        #region 弧度化s
        public static double hudutos(double hudu)
        {
            double d, m, s;
            double du = hudu * 180 / Math.PI;
            d = Math.Floor(du);
            m = Math.Floor((du - d) * 60);
            s = Math.Round(((du - d) * 60 - m) * 60, 1);
            return d * 3600 + m * 60 + s;//保留到0.1秒
        }
        #endregion
        #region 方位角jisuan
        public static double fangwei(double x1, double y1, double x2, double y2)//方位角返回弧度值
        {
            double fangweijiao = 180 - 90 * Math.Abs(y2 - y1) / ((y2 - y1) + Math.Pow(10, -10)) - Math.Atan((x2 - x1) / ((y2 - y1) + Math.Pow(10, -10))) * 180 / Math.PI;
            return fangweijiao * Math.PI / 180;
        }
        #endregion
        #region 绘制三角
        public static void sanjiao(Graphics g, PointF pf)
        {
            //绘制填充多边形的原理
            Bitmap bt1 = new Bitmap(20, 20);//画板
            PointF[] pfs2 = { new PointF(20, 10), new PointF(1, 0), new PointF(1, 20) };//三角的三个点
            Graphics g1 = Graphics.FromImage(bt1);
            g1.FillPolygon(Brushes.White, pfs2);//填充
            g1.DrawPolygon(new Pen(Color.Blue, 1.5f), pfs2);//绘制
            g.DrawImage((Image)bt1, pf.X - 10, pf.Y - 10);//图形绘制的位置
        }
        #endregion
        #region 绘制注记
        public static void ziti(Graphics g, PointF pf, string dianhao)
        {
            Bitmap bt2 = new Bitmap(30, 30);
            Graphics g2 = Graphics.FromImage(bt2);
            g2.RotateTransform(90);
            g2.TranslateTransform(0, -30);//划定原点位置
            g2.DrawString(dianhao, new Font("宋体", 20), Brushes.Green, new Point(5, 5));
            g.DrawImage((Image)bt2, pf.X - 25, pf.Y - 25);
        }
        #endregion
        #region dxf绘制
        public static string shiqian()
        {
            string h;
            h = "0\nSECTION\n2\nTABLES\n0\nTABLE\n2\nLAYER\n0\nLAYER\n70\n0\n2\nshiti\n62\n10\n6\nCONTINUOUS\n";//shiti图层，hongse
            h += "0\nLAYER\n70\n0\n2\nzhuji\n62\n50\n6\nCONTINUOUS\n0\nLAYER\n70\n0\n2\nqita\n62\n90\n6\nCONTINUOUS\n0\nENDTAB\n0\nENDSEC\n0\nSECTION\n";//注记图层，黄色
            return h;
        }
        public static string zhuji(double x1, double y1, string s)
        {
            string h;
            h = "0\nTEXT\n8\nzhuji\n10\n" + x1 + "\n20\n" + y1 + "\n40\n15\n1\n" + s + "\n";
            return h;
        }
        #endregion
    }
}
