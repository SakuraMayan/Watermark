using NPOI.HSSF.UserModel;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HX.MVC.Application.Help.Utilities.File
{
    public class ExcelWatermark
    {

        public static void SetWatermark(string inputfilepath, string waterMarkName, string outputfilepath, int fontEmSize = 16, int xAlign = 50, int yAlign = 80, int rotateC = 40)
        {
            //读出原始文件
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(inputfilepath);


            Font font = new System.Drawing.Font("微软雅黑", fontEmSize);
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //调用DrawText()方法新建图片
                System.Drawing.Image imgWtrmrk = DrawText(waterMarkName, font, System.Drawing.Color.LightGray, System.Drawing.Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth);
               
                sheet.PageSetup.BackgoundImage = imgWtrmrk as Bitmap;
                //sheet.Protect(Guid.NewGuid().ToString(), SheetProtectionType.Content);

            }


            //原始文件复制出来
            workbook.SaveToFile(outputfilepath);
            
            
        }


        public static MemoryStream SetWatermark(Stream stream, string waterMarkName,  int fontEmSize = 16, int xAlign = 50, int yAlign = 80, int rotateC = 40)
        {
            MemoryStream memoryStream = new MemoryStream();

            //读出原始文件
            Workbook workbook = new Workbook();

            try
            {
                workbook.LoadFromStream(stream);


                Font font = new System.Drawing.Font("微软雅黑", fontEmSize);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    //调用DrawText()方法新建图片
                    System.Drawing.Image imgWtrmrk = DrawText(waterMarkName, font, System.Drawing.Color.LightGray, System.Drawing.Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth);

                    sheet.PageSetup.BackgoundImage = imgWtrmrk as Bitmap;
                    //sheet.Protect(Guid.NewGuid().ToString(), SheetProtectionType.Content);

                }


                //原始文件复制出来
                workbook.SaveToStream(memoryStream);
            }
            catch (Exception e)
            {

            }
            finally
            {
               

            }

            

            return memoryStream;

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
                    drawing.DrawString(text, font, textBrush,0,0);
                    drawing.ResetTransform();
                }
            }


            drawing.Save();
            return img;
        }


        /*
         /// <summary>
        /// 删除多余Evaluation Warning页
        /// </summary>
        /// <param name="filePath"></param>
        private string DelSheet(string filePath)
        {
            try
            {
                NPOI.SS.UserModel.IWorkbook workbook = null;
                string tempPath = Path.GetDirectoryName(filePath) + "\\" + "TempDelSheetPath" + Path.GetExtension(filePath);

                FileInfo temp = new FileInfo(tempPath);
                if (temp.Exists)
                {
                    temp.Delete();
                }
                FileStream fs = File.Create(tempPath);
                XSSFWorkbook x1 = new XSSFWorkbook();
                x1.Write(fs);
                fs.Close();
                FileStream fileRead = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (Path.GetExtension(filePath) == ".xls")
                {
                    workbook = new HSSFWorkbook(fileRead);
                    for (int x = 0; x < workbook.NumberOfSheets; x++)
                    {
                        if (workbook.GetSheetName(x) == "Evaluation Warning")
                        {
                            workbook.RemoveSheetAt(x);
                        }
                    }
                }
                else
                {
                    workbook = new XSSFWorkbook(fileRead);
                    int sCounts = workbook.NumberOfSheets;
                    for (int x = 0; x < workbook.NumberOfSheets; x++)
                    {
                        if (workbook.GetSheetName(x) == "Evaluation Warning")
                        {
                            workbook.RemoveSheetAt(x);
                        }
                    }
                }
                fileRead.Close();
                using (FileStream fileSave = new FileStream(tempPath, FileMode.Open, FileAccess.Write))
                {
                    workbook.Write(fileSave);
                }
                workbook.Close();
                return tempPath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
         */


    }

}
