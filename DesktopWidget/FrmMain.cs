using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using DesktopWidget.Properties;

namespace DesktopWidget
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        #region 半透明不规则窗体创建

        private void InitializeStyles()
        {
            SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Selectable, false);
            UpdateStyles();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            InitializeStyles();//设置窗口样式、双缓冲等
            base.OnHandleCreated(e);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00080000; //WS_EX_LAYERED
                return cp;
            }
        }

        public void SetBits(Bitmap bitmap)
        {
            if (!Bitmap.IsCanonicalPixelFormat(bitmap.PixelFormat) || !Bitmap.IsAlphaPixelFormat(bitmap.PixelFormat))
                throw new ApplicationException("The picture must be 32bit picture with alpha channel.");
            IntPtr oldBits = IntPtr.Zero;
            IntPtr screenDC = Win32.GetDC(IntPtr.Zero);
            IntPtr hBitmap = IntPtr.Zero;
            IntPtr memDc = Win32.CreateCompatibleDC(screenDC);
            try
            {
                Win32.Point topLoc = new Win32.Point(Left, Top);
                Win32.Size bitMapSize = new Win32.Size(bitmap.Width, bitmap.Height);
                Win32.BLENDFUNCTION blendFunc = new Win32.BLENDFUNCTION();
                Win32.Point srcLoc = new Win32.Point(0, 0);

                hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
                oldBits = Win32.SelectObject(memDc, hBitmap);

                blendFunc.BlendOp = Win32.AC_SRC_OVER;
                blendFunc.SourceConstantAlpha = 255;
                blendFunc.AlphaFormat = Win32.AC_SRC_ALPHA;
                blendFunc.BlendFlags = 0;

                Win32.UpdateLayeredWindow(Handle, screenDC, ref topLoc, ref bitMapSize, memDc, ref srcLoc, 0, ref blendFunc, Win32.ULW_ALPHA);
            }
            finally
            {
                if (hBitmap != IntPtr.Zero)
                {
                    Win32.SelectObject(memDc, oldBits);
                    Win32.DeleteObject(hBitmap);
                }
                Win32.ReleaseDC(IntPtr.Zero, screenDC);
                Win32.DeleteDC(memDc);
            }
        }
        #endregion

        #region 无边框移动

        private bool isMouseDown = false;
        private Point mouseOffset;

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = true;
                mouseOffset.X = this.Left - Control.MousePosition.X;
                mouseOffset.Y = this.Top - Control.MousePosition.Y;
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            isMouseDown = false;
            base.OnMouseUp(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (isMouseDown)
            {
                var point = Control.MousePosition;
                point.Offset(mouseOffset);
                this.Location = point;
            }
            base.OnMouseMove(e);
        }

        #endregion

        #region 转化为桌面插件

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpWindowClass, string lpWindowName);
        [DllImport("user32.dll")]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
        const int GWL_HWNDPARENT = -8;

        protected override void OnShown(EventArgs e)
        {
            //该事件及设置顺序不可调整
            this.ShowInTaskbar = false;
            SetWindowLong(this.Handle, GWL_HWNDPARENT, GetDesktopPtr());
            base.OnShown(e);
        }

        /// <summary>
        /// 将程序嵌入桌面
        /// </summary>
        /// <returns></returns>
        private IntPtr GetDesktopPtr()//寻找桌面的句柄
        {
            // 情况一
            IntPtr hwndWorkerW = IntPtr.Zero;
            IntPtr hShellDefView = IntPtr.Zero;
            IntPtr hwndDesktop = IntPtr.Zero;
            IntPtr hProgMan = FindWindow("Progman", "Program Manager");
            if (hProgMan != IntPtr.Zero)
            {
                hShellDefView = FindWindowEx(hProgMan, IntPtr.Zero, "SHELLDLL_DefView", null);
                if (hShellDefView != IntPtr.Zero)
                    hwndDesktop = FindWindowEx(hShellDefView, IntPtr.Zero, "SysListView32", null);
            }
            if (hwndDesktop != IntPtr.Zero)
                return hwndDesktop;
            // 情况二
            //存在桌面窗口层次
            while (hwndDesktop == IntPtr.Zero)
            {
                //获得WorkerW类的窗口
                hwndWorkerW = FindWindowEx(IntPtr.Zero, hwndWorkerW, "WorkerW", null);
                if (hwndWorkerW == IntPtr.Zero)
                    break;
                hShellDefView = FindWindowEx(hwndWorkerW, IntPtr.Zero, "SHELLDLL_DefView", null);
                if (hShellDefView == IntPtr.Zero)
                    continue;
                hwndDesktop = FindWindowEx(hShellDefView, IntPtr.Zero, "SysListView32", null);
            }
            return hwndDesktop;
        }

        #endregion

        SolidBrush brush;
        Bitmap bmp, bmpBg;
        Graphics gBmp;
        //字体样式须为局部变量
        PrivateFontCollection pfc;
        JavaScriptSerializer jsonConvert = new JavaScriptSerializer();
        Font font;
        string configPath = "";
        //xi'an = 101110101
        //beijing = 101010100
        //shanghai = 101020100 
        Config config = new Config() { CityCode = "101110101" };
        Dictionary<string, string> dicWeather = new Dictionary<string, string>()
        {
            { "weather","--"},
            { "temp","--"},
            { "wind","--"},
            { "aqi","--"}
        };

        private void FrmMain_Load(object sender, EventArgs e)
        {
            brush = new SolidBrush(Color.FromArgb(147, 174, 97));
            //直接调用Resource对象会导致内存开销变大
            bmpBg = Resources.glass;
            bmp = new Bitmap(this.Width, bmpBg.Height * this.Width / bmpBg.Width);
            gBmp = Graphics.FromImage(bmp);
            pfc = new PrivateFontCollection();
            var fontData = Resources.Font;
            IntPtr iFont = Marshal.AllocHGlobal(fontData.Length);
            Marshal.Copy(fontData, 0, iFont, fontData.Length);
            pfc.AddMemoryFont(iFont, fontData.Length);
            Marshal.FreeHGlobal(iFont);
            font = new Font(pfc.Families[0], 14f);
            configPath = Application.StartupPath + "\\dw.dat";
            if (File.Exists(configPath))
            {
                config = jsonConvert.Deserialize<Config>(File.ReadAllText(configPath));
                this.Location = config.Location;
            }
            new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        GetWeather();
                        Thread.Sleep(TimeSpan.FromHours(1));
                    }
                    catch
                    {
                        Thread.Sleep(TimeSpan.FromMinutes(1));
                    }
                }
            })
            { IsBackground = true }.Start();
        }

        private void FrmMain_Shown(object sender, EventArgs e)
        {
            this.Draw();
            this.timer.Start();
        }

        protected override void OnLocationChanged(EventArgs e)
        {
            if (string.IsNullOrEmpty(this.configPath))
                return;
            this.config.Location = this.Location;
            File.WriteAllText(configPath, jsonConvert.Serialize(this.config));
        }

        private void GetWeather()
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create($"http://d1.weather.com.cn/sk_2d/{config.CityCode}.html?_={DateTime.Now.ToLinuxTime()}");
            httpWebRequest.Referer = $"http://en.weather.com.cn/weather/{config.CityCode}.shtml";
            httpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36";
            using (HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse())
            {
                using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
                {
                    var strWeather = streamReader.ReadToEnd();
                    var weather = jsonConvert.Deserialize<Dictionary<string, string>>(strWeather.Substring(strWeather.IndexOf('{')));
                    dicWeather["weather"] = weather["weathere"].Substring(0, 1).ToUpper() + weather["weathere"].Substring(1);
                    dicWeather["temp"] = weather["temp"];
                    dicWeather["wind"] = weather["wde"] + weather["wse"].Replace("&lt;", "<").Replace("&gt;", ">");
                    dicWeather["aqi"] = weather["aqi"];
                    httpWebResponse.Close();
                    streamReader.Close();
                };
            }
        }

        void Draw()
        {
            gBmp.Clear(Color.Transparent);
            gBmp.FillRectangle(brush, 25, 25, 230, 160);
            string strText = string.Format($@"
   {DateTime.Now.ToString("MM-dd HH:mm:ss")}
 Weather {dicWeather["weather"]}
    Temp {dicWeather["temp"]}`C
    Wind {dicWeather["wind"]}
     AQI {dicWeather["aqi"]}
");
            int y = -10;
            foreach (var text in strText.Split('\r'))
            {
                this.gBmp.DrawString(text, this.font, Brushes.Black, 25, y);
                y += 30;
            }
            gBmp.DrawImage(bmpBg, 0, 0, bmp.Width, bmp.Height);
            SetBits(bmp);
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            Draw();
        }
    }

    public class Config
    {
        public Point Location { get; set; }

        public string CityCode { get; set; }
    }

    public static class Extension
    {
        public static int ToLinuxTime(this DateTime time)
        {
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            return (int)(time - startTime).TotalSeconds;
        }
    }
}
