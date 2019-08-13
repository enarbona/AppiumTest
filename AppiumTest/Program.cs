using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;

namespace AppiumTest
{
    class Program
    {
        private  static Process windowsMailProcess;
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
       private  static List<Bitmap> captureClip=new List<Bitmap>();
        private static WindowsDriver<WindowsElement> cortanaSession;
        private static WindowsDriver<WindowsElement> desktopSession;
        static void Main(string[] args)
        {
            windowsMailProcess = new Process();

            windowsMailProcess.StartInfo.UseShellExecute = true;

            windowsMailProcess.StartInfo.FileName = "2019-04-19_133828_114229_web.eml";

            windowsMailProcess.Start();


            Thread.Sleep(1500);

            DesiredCapabilities desktopCapabilities = new DesiredCapabilities();
            desktopCapabilities.SetCapability("app", "Root");
            desktopSession = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), desktopCapabilities);
            WindowsElement winMail = desktopSession.FindElementByName("Mail");


            var document = winMail.FindElementByName("Message");


            var scrollable = bool.Parse(document.GetAttribute("Scroll.VerticallyScrollable"));

            if (!scrollable)
            {

                document.GetScreenshot().SaveAsFile(@"c:\archine\appiumTest.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            else {

                Actions action = new Actions(desktopSession);
              

                var ScrollVerticalViewSize = double.Parse(document.GetAttribute("Scroll.VerticalViewSize"));

                var ScrollVerticalScrollPercen = double.Parse(document.GetAttribute("Scroll.VerticalScrollPercent"));
                var ParseBoundingRectangle = JObject.Parse("{" +document.GetAttribute("BoundingRectangle").Replace(' ',',')+"}");
                var BoundingRectangle = new Rectangle(
                    ParseBoundingRectangle["Left"].Value<int>(),
                    ParseBoundingRectangle["Top"].Value<int>(),
                    ParseBoundingRectangle["Width"].Value<int>(),
                    ParseBoundingRectangle["Height"].Value<int>());


                var totalArea = Convert.ToInt32(System.Math.Floor((BoundingRectangle.Height * 100) / ScrollVerticalViewSize));


                var steeps = Math.Floor((double)totalArea / (double)BoundingRectangle.Height);
                action.MoveByOffset(BoundingRectangle.X+5, BoundingRectangle.Y + 5).Perform();
                for (int i = 0; i < steeps; i++)
                {
                    captureClip.Add(ArrayToBitmap(document.GetScreenshot().AsByteArray));
                    document.SendKeys(Keys.PageDown);
                    Thread.Sleep(100);
                   
                }

                MergedBitmaps(captureClip).Save(@"c:\archive\appiumTest2.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

                captureClip.Clear();
            }
        }
        private static Bitmap ArrayToBitmap(byte[] bytes)
        {



            using (var ms = new MemoryStream(bytes))
            {
                return new Bitmap(ms);
            }

        }
        private static Bitmap MergedBitmaps(List<Bitmap> bitmaps)
        {
            int maxHeight = bitmaps.Sum(r => r.Height - 1);
            int maxWidth = bitmaps.Max(r => r.Width);
            Bitmap result = new Bitmap(maxWidth, maxHeight);

            using (Graphics g = Graphics.FromImage(result))
            {
                int y = 0;
                int x = 0;
                foreach (Bitmap tmp in bitmaps)
                {
                    g.DrawImage(tmp, new Point(x, y));
                    y = y + tmp.Height - 1;
                    tmp.Dispose();
                }
            }


            if (result.Height > 7000)
            {

                Rectangle cropArea = new Rectangle(new Point(0, 0), new Size(result.Width, 7000));
              
                result = result.Clone(cropArea, result.PixelFormat);
            }

            GC.Collect();


            return result;

        }
    }
}
