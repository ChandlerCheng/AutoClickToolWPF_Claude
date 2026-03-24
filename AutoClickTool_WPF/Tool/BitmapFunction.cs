using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoClickTool_WPF.Tool
{
    public class BitmapFunction
    {
        public static double CalculateColorRatio(Bitmap bitmap, System.Drawing.Color targetColor)
        {
            // 獲取圖像的寬度和高度
            int width = bitmap.Width;
            int height = bitmap.Height;

            // 初始化目標顏色像素數量
            int targetColorCount = 0;

            // 遍歷圖像的每個像素，計算目標顏色的像素數量
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    // 取得當前像素的顏色
                    System.Drawing.Color pixelColor = bitmap.GetPixel(x, y);

                    // 檢查像素是否與目標顏色匹配
                    if (pixelColor.R == targetColor.R && pixelColor.G == targetColor.G && pixelColor.B == targetColor.B)
                    {
                        targetColorCount++;
                    }
                }
            }

            // 計算目標顏色的佔據比例
            double targetColorRatio = (double)targetColorCount / (width * height);
            return Math.Round(targetColorRatio, 2);
        }
        public static Bitmap CaptureScreen(int x, int y, int width, int height)
        {
            try
            {
                // 建立 Bitmap，格式為 24bppRgb
                Bitmap bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                // 使用 Graphics 擷取指定範圍的螢幕畫面
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                }

                return bitmap;  // 返回擷取的 Bitmap
            }
            catch (Exception ex)
            {
                return null;  // 如果發生錯誤，返回 null
            }
        }
        public static double CompareImages(Bitmap image1, Bitmap image2)
        {
            if (image1.Width != image2.Width || image1.Height != image2.Height)
                throw new ArgumentException("圖像大小不一致");

            int width = image1.Width;
            int height = image1.Height;

            // 計算相似度
            double difference = 0;
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    System.Drawing.Color pixel1 = image1.GetPixel(x, y);
                    System.Drawing.Color pixel2 = image2.GetPixel(x, y);

                    // 比較 RGB 差異，允許一定範圍內的容差
                    if (Math.Abs(pixel1.R - pixel2.R) > 10 || Math.Abs(pixel1.G - pixel2.G) > 10 || Math.Abs(pixel1.B - pixel2.B) > 10)
                    {
                        difference++;
                    }
                }
            }

            double totalPixels = width * height;
            double percentageSimilarity = (totalPixels - difference) / totalPixels * 100;

            return percentageSimilarity;
        }

        public static bool CompareGameScreenshots(Bitmap compareBmp, int captureX, int captureY, double ratio)
        {
            int xOffset = Coordinate.windowBoxLineOffset;
            int yOffset = Coordinate.windowHOffset;
            int x = captureX + xOffset;
            int y = captureY + yOffset;
            int width = compareBmp.Width;
            int height = compareBmp.Height;

            Bitmap screenshot_Bmp = CaptureScreen(x, y, width, height);
            double Final_NCP = CompareImages(screenshot_Bmp, compareBmp);

            //MessageBox.Show($"比對值 '{Final_NCP}' \n");

            if (Final_NCP == (double)100)
                return true;
            else if (Final_NCP > ratio)
                return true;
            else
                return false;
        }
    }
}
