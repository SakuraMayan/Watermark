using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace HX.MVC.Application.Help.Utilities.File
{
    public class ConvertPDF
    {
        private static string fontPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"font\msyh.ttf");
        //private static string fontPath = "https://s3.cn-northwest-1.amazonaws.com.cn/hx-fxzl/font/SIMFANG.TTF";
        /// <summary>
        /// 添加普通偏转角度文字水印保存到文件
        /// </summary>
        /// <param name="inputfilepath">文件地址</param>
        /// <param name="outputfilepath"></param>
        /// <param name="waterMarkName"></param>
        public static void SetWatermarkToPath(string inputfilepath, string waterMarkName, string outputfilepath, string fontS = "16", string bigWaterMarkName = "",int xAlign=50,int yAlign=80,int rotateC=40)
        {
            //if (string.IsNullOrEmpty(waterMarkName))
            //{
            //    try { File.Delete(inputfilepath); } catch (Exception e) { }
            //    throw new Exception("水印文字不能为空");
            //}

            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            FileStream stream = null;
            try
            {
                stream = new FileStream(outputfilepath, FileMode.Create);
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, stream);
                pdfStamper.SetEncryption(PdfWriter.STRENGTH40BITS, null, Guid.NewGuid().ToString(), PdfWriter.AllowScreenReaders);
                int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState()
                {
                    FillOpacity = 0.2f//透明度
                };
                //for (int i = 1; i < total; i++)
                //{
                //    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                //    content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                //    //透明度
                //    gs.FillOpacity = 0.3f;
                //    content.SetGState(gs);
                //    content.SetGrayFill(0.3f);
                //    //开始写入文本
                //    content.BeginText();
                //    content.SetColorFill(BaseColor.LightGray);
                //    content.SetFontAndSize(font, 100);
                //    content.SetTextMatrix(0, 0);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 - 50, height / 2 - 50, 55);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 +50, height / 2+50, 55);
                //    //content.SetColorFill(BaseColor.BLACK);
                //    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                //    content.EndText();
                //}
                int span = 1;//设置垂直位移
                int spanHeight = yAlign;//垂直间距 垂直位移间距
                int fontSize = !string.IsNullOrEmpty(fontS) ? Convert.ToInt32(fontS) : 25;//设置字体大小
                int waterMarkNameLenth = waterMarkName.Length;
                char c;
                int rise = 0;
                string spanString = "";//字符间距
                int spanWidth = 1;//水平位移
                int spanWaterMarkNameWidth = xAlign;//水印文字位移宽度 水平位移宽度
                for (int i = 1; i < total; i++)
                {
                    rise = waterMarkNameLenth * span+ spanHeight;
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    content.SetGState(gs);
                    content.BeginText();
                    //content.SetColorFill(Color.LIGHT_GRAY);
                    content.SetFontAndSize(font, fontSize);
                    

                    int heightNumbert = (int)Math.Ceiling((decimal)height / (decimal)rise);//垂直重复的次数，进一发
                    int panleWith = fontSize * waterMarkNameLenth/2+spanWaterMarkNameWidth;
                    int widthNumber = (int)Math.Ceiling((decimal)width / (decimal)panleWith);//水平重复次数

                    for (int w = 0; w < widthNumber; w++)
                    {

                        for (int h = 1, f = heightNumbert-1; h <= heightNumbert;h++,f--)
                        {

                            int yleng = rise * h;
                            int wleng = (w * panleWith) + (spanWidth * f);
                            
                            content.SetTextMatrix(wleng, yleng);//x,y设置水印开始的绝对左边，以左下角为x，y轴的起点

                            content.ShowTextAligned(Element.ALIGN_JUSTIFIED_ALL, waterMarkName, wleng, yleng, rotateC);//控制倾斜角度

                            //content.ShowText(spanString);
                            //for (int k = 0; k < waterMarkNameLenth; k++)
                            //{
                            //    content.SetTextRise(yleng);//指定的y轴值处添加
                            //    c = waterMarkName[k];
                            //    content.ShowText(c + spanString);
                            //    yleng += span;
                            //}
                        }
                    }

                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();


                if (pdfReader != null)
                    pdfReader.Close();

            }
        }


        /// <summary>
        /// 添加普通偏转角度文字水印保存到文件
        /// </summary>
        /// <param name="isp">文件流</param>
        /// <param name="outputfilepath"></param>
        /// <param name="waterMarkName"></param>
        /// <param name="permission"></param>
        public static void SetWatermarkToPath(Stream isp, string waterMarkName, string outputfilepath, string fontS = "16", string bigWaterMarkName="", int xAlign = 50, int yAlign = 80, int rotateC = 40)
        {
            //if (string.IsNullOrEmpty(waterMarkName))
            //{
            //    return;
            //}

            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(isp);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                pdfStamper.SetEncryption(PdfWriter.STRENGTH40BITS, null, Guid.NewGuid().ToString(), PdfWriter.AllowScreenReaders);
                int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState()
                {
                    FillOpacity = 0.2f//透明度
                };
                //for (int i = 1; i < total; i++)
                //{
                //    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                //    content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                //    //透明度
                //    gs.FillOpacity = 0.3f;
                //    content.SetGState(gs);
                //    content.SetGrayFill(0.3f);
                //    //开始写入文本
                //    content.BeginText();
                //    content.SetColorFill(BaseColor.LightGray);
                //    content.SetFontAndSize(font, 100);
                //    content.SetTextMatrix(0, 0);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 - 50, height / 2 - 50, 55);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 +50, height / 2+50, 55);
                //    //content.SetColorFill(BaseColor.BLACK);
                //    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                //    content.EndText();
                //}
                int span = 1;//设置垂直位移
                int spanHeight = yAlign;//垂直间距
                int fontSize = !string.IsNullOrEmpty(fontS) ? Convert.ToInt32(fontS) : 25;//设置字体大小
                int waterMarkNameLenth = waterMarkName.Length;
                char c;
                int rise = 0;
                string spanString = "";//字符间距
                int spanWidth = 1;//水平位移
                int spanWaterMarkNameWidth = xAlign;//水印文字位移宽度
                for (int i = 1; i < total; i++)
                {
                    rise = waterMarkNameLenth * span + spanHeight;
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    content.SetGState(gs);
                    content.BeginText();
                    //content.SetColorFill(Color.LIGHT_GRAY);
                    content.SetFontAndSize(font, fontSize);


                    int heightNumbert = (int)Math.Ceiling((decimal)height / (decimal)rise);//垂直重复的次数，进一发
                    int panleWith = fontSize * waterMarkNameLenth / 2 + spanWaterMarkNameWidth;
                    int widthNumber = (int)Math.Ceiling((decimal)width / (decimal)panleWith);//水平重复次数

                    for (int w = 0; w < widthNumber; w++)
                    {
                        for (int h = 1, f = heightNumbert - 1; h <= heightNumbert; h++, f--)
                        {
                            int yleng = rise * h;
                            int wleng = (w * panleWith) + (spanWidth * f);
                            content.SetTextMatrix(wleng, yleng);//x,y设置水印开始的绝对左边，以左下角为x，y轴的起点

                            content.ShowTextAligned(Element.ALIGN_JUSTIFIED_ALL, waterMarkName, wleng, yleng, rotateC);

                            //content.ShowText(spanString);
                            //for (int k = 0; k < waterMarkNameLenth; k++)
                            //{
                            //    content.SetTextRise(yleng);//指定的y轴值处添加
                            //    c = waterMarkName[k];
                            //    content.ShowText(c + spanString);
                            //    yleng += span;
                            //}
                        }
                    }

                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

        /// <summary>
        /// 添加普通偏转角度文字水印（保存到内存）
        /// </summary>
        /// <param name="inStream"></param>
        /// <param name="waterMarkName"></param>
        /// <param name="fontS"></param>
        /// <param name="bigWaterMarkName"></param>
        /// <returns></returns>
        public static MemoryStream SetWatermark(Stream inStream, string waterMarkName,string fontS= "16", string bigWaterMarkName = "", int xAlign = 50, int yAlign = 80, int rotateC = 40)
        {
            MemoryStream outStream = new MemoryStream();

            if (string.IsNullOrEmpty(waterMarkName))
            {
                inStream.CopyTo(outStream);
                return outStream;
            }


            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inStream);
                pdfStamper = new PdfStamper(pdfReader, outStream);
                pdfStamper.SetEncryption(PdfWriter.STRENGTH40BITS, null, Guid.NewGuid().ToString(), PdfWriter.AllowScreenReaders);
                int pages = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState()
                {
                    FillOpacity = 0.2f//透明度
                };
                //for (int i = 1; i < total; i++)
                //{
                //    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                //    content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                //    //透明度
                //    gs.FillOpacity = 0.3f;
                //    content.SetGState(gs);
                //    content.SetGrayFill(0.3f);
                //    //开始写入文本
                //    content.BeginText();
                //    content.SetColorFill(BaseColor.LightGray);
                //    content.SetFontAndSize(font, 100);
                //    content.SetTextMatrix(0, 0);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 - 50, height / 2 - 50, 55);
                //    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 +50, height / 2+50, 55);
                //    //content.SetColorFill(BaseColor.BLACK);
                //    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                //    content.EndText();
                //}
                int span = 1;//设置垂直位移
                int spanHeight = yAlign;//垂直间距
                int fontSize = !string.IsNullOrEmpty(fontS) ? Convert.ToInt32(fontS) : 25;//设置字体大小
                int waterMarkNameLenth = waterMarkName.Length;
                char c;
                int rise = 0;
                string spanString = "";//字符间距
                int spanWidth = 1;//水平位移
                int spanWaterMarkNameWidth = xAlign;//水印文字位移宽度
                for (int i = 1; i < pages; i++)
                {
                    rise = waterMarkNameLenth * span+ spanHeight;
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    content.SetGState(gs);
                    content.BeginText();
                    //content.SetColorFill(Color.LIGHT_GRAY);
                    content.SetFontAndSize(font, fontSize);


                    int heightNumbert = (int)Math.Ceiling((decimal)height / (decimal)rise);//垂直重复的次数，进一发
                    int panleWith = fontSize * waterMarkNameLenth/2+spanWaterMarkNameWidth;
                    int widthNumber = (int)Math.Ceiling((decimal)width / (decimal)panleWith);//水平重复次数

                    for (int w = 0; w < widthNumber; w++)
                    {
                        for (int h = 1, f = heightNumbert-1; h <= heightNumbert;h++,f--)
                        {
                            int yleng = rise * h;
                            int wleng = (w * panleWith) + (spanWidth * f);
                            content.SetTextMatrix(wleng, yleng);//x,y设置水印开始的绝对左边，以左下角为x，y轴的起点

                            content.ShowTextAligned(Element.ALIGN_JUSTIFIED_ALL, waterMarkName, wleng, yleng, rotateC);

                            //content.ShowText(spanString);
                            //for (int k = 0; k < waterMarkNameLenth; k++)
                            //{
                            //    content.SetTextRise(yleng);//指定的y轴值处添加
                            //    c = waterMarkName[k];
                            //    content.ShowText(c + spanString);
                            //    yleng += span;
                            //}
                        }
                    }

                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                 
                

                throw ex;
            }
            finally
            {
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }

            return outStream;
        }


        //public static string GetMapPath(string strPath)
        //{
        //    if (HttpContext.Current != null)
        //    {
        //        return HttpContext.Current.Server.MapPath(strPath);
        //    }
        //    else //非web程序引用
        //    {
        //        return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, strPath);
        //    }
        //}
    }
}
