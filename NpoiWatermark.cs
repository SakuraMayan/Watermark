
using NPOI;
using NPOI.OpenXml4Net.OPC;
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
    public class NpoiWatermark
    {

        /*public static BufferedImage createWatermarkImage(Watermark watermark)
        {
            if (watermark == null)
            {
                watermark = new Watermark();
                watermark.enable=true;
                //            watermark.setText("userName");
                watermark.text="内部资料";
                watermark.color="#C5CBCF";
                watermark.dateFormat="yyyy-MM-dd HH:mm";
            }
            *//*else
            {
                if (StringUtils.isEmpty(watermark.getDateFormat()))
                {
                    watermark.setDateFormat("yyyy-MM-dd HH:mm");
                }
                else if (watermark.getDateFormat().length() == 16)
                {
                    watermark.setDateFormat("yyyy-MM-dd HH:mm");
                }
                else if (watermark.getDateFormat().length() == 10)
                {
                    watermark.setDateFormat("yyyy-MM-dd");
                }
                if (StringUtils.isEmpty(watermark.getText()))
                {
                    watermark.setText("内部资料");
                }
                if (StringUtils.isEmpty(watermark.getColor()))
                {
                    watermark.setColor("#C5CBCF");
                }
            }*//*
            String[] textArray = watermark.text.Split("\n");
            Font font = new Font("microsoft-yahei", Font.PLAIN, 20);
            int width = 300;
            int height = 100;

            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            // 背景透明 开始
            Graphics2D g = image.createGraphics();
            image = g.getDeviceConfiguration().createCompatibleImage(width, height, Transparency.TRANSLUCENT);
            g.dispose();
            // 背景透明 结束
            g = image.createGraphics();
            g.setColor(Color.LIGHT_GRAY);// 设定画笔颜色
            g.setFont(font);// 设置画笔字体
            g.shear(0.1, -0.26);// 设定倾斜度

            //        设置字体平滑
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

            int y = 50;
            for (int i = 0; i < textArray.Length; i++)
            {
                g.drawString(textArray[i], 0, y);// 画出字符串
                y = y + font.getSize();
            }
            //g.drawString(DateUtils.getNowDateFormatCustom(watermark.getDateFormat()), 0, y);// 画出字符串

            g.dispose();// 释放画笔
            return image;

        }*/

        public class Watermark
        {
            public Boolean enable { get; set; }
            public String text { get; set; }
            public String dateFormat { get; set; }
            public String color { get; set; }
        }


        public static void setWatermark(string inputfilepath, string waterMarkName, string outputfilepath)
        {

            try
            {
                FileStream fs = new FileStream(outputfilepath,FileMode.Create);

                XSSFWorkbook workbook = new XSSFWorkbook(inputfilepath);


                //add picture data to this workbook.
                //                FileInputStream is = new FileInputStream("/Users/Tony/Downloads/data_image.png");
                //            byte[] bytes = IOUtils.toByteArray(is);

                Font font = new System.Drawing.Font("微软雅黑", 16);

                Image image = DrawText(waterMarkName, font, System.Drawing.Color.LightGray, System.Drawing.Color.White, 100, 100);
                // 导出到字节流B
                MemoryStream os = new MemoryStream();
                image.Save(os, System.Drawing.Imaging.ImageFormat.Png);

                int pictureIdx = workbook.AddPicture(os.ToArray(), PictureType.PNG);
                POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart)workbook.GetAllPictures()[pictureIdx];
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {//获取每个Sheet表
                    XSSFSheet sheetWB = (XSSFSheet)workbook.GetSheetAt(i);
                    PackagePartName ppn = poixmlDocumentPart.GetPackagePart().PartName;
                    String relType = XSSFRelation.IMAGES.Relation;
                    //add relation from sheet to the picture data
                    PackageRelationship pr = sheetWB.GetPackagePart().AddRelationship(ppn, TargetMode.Internal, relType, null);
                    //set background picture to sheet
                    sheetWB.GetCTWorksheet().picture = new NPOI.OpenXmlFormats.Spreadsheet.CT_SheetBackgroundPicture() {
                        id = pr.Id
                    };
                }

                workbook.Write(fs);


            }
            catch (Exception e)
            {
               
            }

        }

        private static System.Drawing.Image DrawText(String text, System.Drawing.Font font, Color textColor, Color backColor, double height, double width)
        {
            //创建一个指定宽度和高度的位图图像
            Image img = new Bitmap((int)width, (int)height);
            Graphics drawing = Graphics.FromImage(img);
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
