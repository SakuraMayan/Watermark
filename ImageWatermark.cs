using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace HX.MVC.Application.Help.Utilities.File
{
    public class ImageWatermark
    {
        /// <summary>
        /// 图片水印
        /// </summary>
        /// <param name="s">S3的文件流</param>
        /// <param name="waterMarkName">水印名</param>
        /// <returns></returns>
        public static MemoryStream AddImageText(Stream s, string waterMarkName, int fontEmSize = 16, int xAlign = 90, int yAlign = 120, int rotateC = -40)
        {
            MemoryStream outStream = new MemoryStream();

            if (string.IsNullOrEmpty(waterMarkName))
            {
                s.CopyTo(outStream);
                return outStream;
            }

            Image image = Image.FromStream(s);
            // 获取图片的宽和高
            int width = image.Width;
            int height = image.Height;
            Graphics g = Graphics.FromImage(image);
            // 字体 字号 字体样式
            Font drawFont = new Font("微软雅黑", fontEmSize, FontStyle.Italic, GraphicsUnit.Pixel);
            SizeF fontSize = g.MeasureString(waterMarkName, drawFont);
            float xpos = 0;
            float ypos = 0;
            // 倾斜 -30°
            // g.RotateTransform(-2);
            // 循环打印文字水印
            for (int i = -width / 4; i < width * 3; i += xAlign)
            {
                xpos = i;
                for (int j = -height / 4; j < height * 3; j += yAlign)
                {
                    ypos = j;
                    //g.DrawString(waterMarkName, drawFont, new SolidBrush(Color.FromArgb(50, 100, 100, 100)), xpos, ypos);
                    g.TranslateTransform(xpos, ypos);
                    g.RotateTransform(rotateC);
                    g.DrawString(waterMarkName, drawFont, new SolidBrush(Color.FromArgb(50, 100, 100, 100)), 0, 0);
                    g.ResetTransform();
                }
            }

            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            ImageCodecInfo ici = null;
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.MimeType.IndexOf("jpeg") > -1)
                    ici = codec;
            }
            EncoderParameters encoderParams = new EncoderParameters();
            long[] qualityParam = new long[1];

            // 图片质量：0 - 100
            qualityParam[0] = 100;

            EncoderParameter encoderParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qualityParam);
            encoderParams.Param[0] = encoderParam;
            image.Save(outStream, ici, encoderParams);
            g.Dispose();
            image.Dispose();

            return outStream;

        }

        /// <summary>
        /// 图片水印
        /// </summary>
        /// <param name="inputfilepath">文件地址</param>
        /// <param name="waterMarkName">水印名</param>
        /// <returns></returns>
        public static void AddImageText(string inputfilepath, string waterMarkName, string outputfilepath,int fontEmSize=16,int xAlign=90,int yAlign=120,int rotateC= -40)
        {
            if (string.IsNullOrEmpty(waterMarkName)||string.IsNullOrEmpty(inputfilepath))
            {
                return;
            }

            using (var s = new FileStream(inputfilepath, FileMode.Open, FileAccess.Read))
            {
                Image image = Image.FromStream(s);
                // 获取图片的宽和高
                int width = image.Width;
                int height = image.Height;
                Graphics g = Graphics.FromImage(image);
                // 字体 字号 字体样式
                Font drawFont = new Font("微软雅黑", fontEmSize, FontStyle.Italic, GraphicsUnit.Pixel);
                SizeF fontSize = g.MeasureString(waterMarkName, drawFont);
                float xpos = 0;
                float ypos = 0;
                // 倾斜 -30°
                //g.RotateTransform(-2);
                // 循环打印文字水印
                for (int i = -width / 4; i < width * 3; i += xAlign)//控制水印水平间隔
                {
                    xpos = i;
                    for (int j = -height / 4; j < height * 3; j += yAlign)//控制水印垂直间隔
                    {
                        ypos = j;
                        //g.DrawString(waterMarkName, drawFont, new SolidBrush(Color.FromArgb(50, 100, 100, 100)), xpos, ypos);
                        g.TranslateTransform(xpos, ypos);
                        g.RotateTransform(rotateC);//控制倾斜角度
                        g.DrawString(waterMarkName, drawFont, new SolidBrush(Color.FromArgb(50, 100, 100, 100)), 0, 0);
                        g.ResetTransform();
                    }
                }

                ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
                ImageCodecInfo ici = null;
                foreach (ImageCodecInfo codec in codecs)
                {
                    if (codec.MimeType.IndexOf("jpeg") > -1)
                        ici = codec;
                }
                EncoderParameters encoderParams = new EncoderParameters();
                long[] qualityParam = new long[1];

                // 图片质量：0 - 100
                qualityParam[0] = 100;

                EncoderParameter encoderParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qualityParam);
                encoderParams.Param[0] = encoderParam;
                image.Save(outputfilepath, ici, encoderParams);
                g.Dispose();
                image.Dispose();
            }

        }
    }
}
