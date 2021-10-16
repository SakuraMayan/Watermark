using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestEquals
{
    public class NpoiWatermark2
    {

        public static void setWatermark(string inputfilepath, string waterMarkName, string outputfilepath)
        {
            FileStream fs = new FileStream(outputfilepath, FileMode.Create);

            XSSFWorkbook workbook = new XSSFWorkbook(inputfilepath);
            Font font = new System.Drawing.Font("微软雅黑", 16);

            Image image = DrawText(waterMarkName, font, System.Drawing.Color.LightGray, System.Drawing.Color.White, 100, 100);

            putWaterRemarkToExcel(workbook, workbook.GetSheetAt(0) as XSSFSheet, image, 0, 0, 0, 10, 1, 3, 0, 0);//生成水印

            workbook.Write(fs);
        }

        public static void putWaterRemarkToExcel(XSSFWorkbook workbook, XSSFSheet sheet, Image image, int startXCol,
                int startYRow, int betweenXCol, int betweenYRow, int XCount, int YCount, int waterRemarkWidth,
                int waterRemarkHeight)
        {



            MemoryStream os = new MemoryStream();
            image.Save(os, System.Drawing.Imaging.ImageFormat.Png);

            // 开始打水印
            var drawing = sheet.CreateDrawingPatriarch();

            // 按照共需打印多少行水印进行循环
            for (int yCount = 0; yCount < YCount; yCount++)
            {
                // 按照每行需要打印多少个水印进行循环
                for (int xCount = 0; xCount < XCount; xCount++)
                {
                    // 创建水印图片位置
                    int xIndexInteger = startXCol + (xCount * waterRemarkWidth) + (xCount * betweenXCol);
                    int yIndexInteger = startYRow + (yCount * waterRemarkHeight) + (yCount * betweenYRow);
                    /*
                     * 参数定义： 第一个参数是（x轴的开始节点）； 第二个参数是（是y轴的开始节点）； 第三个参数是（是x轴的结束节点）；
                     * 第四个参数是（是y轴的结束节点）； 第五个参数是（是从Excel的第几列开始插入图片，从0开始计数）；
                     * 第六个参数是（是从excel的第几行开始插入图片，从0开始计数）； 第七个参数是（图片宽度，共多少列）；
                     * 第8个参数是（图片高度，共多少行）；
                     */
                    IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, xIndexInteger,
                            yIndexInteger, xIndexInteger + waterRemarkWidth, yIndexInteger + waterRemarkHeight);
                    
                    var pic = drawing.CreatePicture(anchor,
                            workbook.AddPicture(os.ToArray(), PictureType.PNG));
                    pic.Resize();
                }
            }
        }

        private static System.Drawing.Image DrawText(String text, System.Drawing.Font font, Color textColor, Color backColor, double height, double width)
        {
            //创建一个指定宽度和高度的位图图像
            Image img = new Bitmap((int)width, (int)height);
            Graphics drawing = Graphics.FromImage(img);
            drawing.Clear(Color.Transparent);//创建透明图片
            //获取文本大小
            SizeF textSize = drawing.MeasureString(text, font);
            //旋转图片
            //drawing.TranslateTransform(((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);
            //drawing.RotateTransform(-45);
            //drawing.TranslateTransform(-((int)width - textSize.Width) / 2, -((int)height - textSize.Height) / 2);
            //绘制背景
            drawing.Clear(backColor);
            //创建文本刷
            Brush textBrush = new SolidBrush(textColor);

            float xpos = 0;
            float ypos = 0;
            // 倾斜 -30°
            //g.RotateTransform(-2);
            // 循环打印文字水印
            for (int i = -(int)width / 4; i < width * 3; i += 90)//控制水印水平间隔
            {
                xpos = i;
                for (int j = -(int)height / 4; j < height * 3; j += 120)//控制水印垂直间隔
                {
                    ypos = j;
                    //g.DrawString(waterMarkName, drawFont, new SolidBrush(Color.FromArgb(50, 100, 100, 100)), xpos, ypos);
                    drawing.TranslateTransform(xpos, ypos);
                    drawing.RotateTransform(-45);//控制倾斜角度
                    drawing.DrawString(text, font, textBrush, 0, 0);
                    drawing.ResetTransform();
                }
            }


            drawing.Save();
            return img;
        }
    }

}
