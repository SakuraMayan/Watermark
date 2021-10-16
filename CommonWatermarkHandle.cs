using HX.AccessControl.SYS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace HX.MVC.Application.Help.Utilities.File
{
    /// <summary>
    /// 公共的水印处理
    /// </summary>
    public class CommonWatermarkHandle
    {
        public static string[] ImageExtensions = { "bmp", "png", "jpeg", "jpg" };
        public static string[] PdfExtensions = { "pdf" };
        public static string[] ExcelExtensions = { "xls", "xlsx" };

        /// <summary>
        /// 流过水印
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="sysUser">用户信息</param>
        /// <param name="stream">流信息</param>
        /// <returns></returns>
        public static MemoryStream AddWatermark(string fileName, string waterMark, Stream stream)
        {

            //判断文件后缀
            var splitFileNames = fileName.Split('.');
            Array.Reverse(splitFileNames);
            string fileExtension = splitFileNames?[0];

            MemoryStream ms = new MemoryStream();


            if (ImageExtensions.Any(t => t.Equals(fileExtension, StringComparison.OrdinalIgnoreCase)))
            { //image
                ms = ImageWatermark.AddImageText(stream, waterMark);
            }
            else if (PdfExtensions.Any(t => t.Equals(fileExtension, StringComparison.OrdinalIgnoreCase)))
            {//pdf 

                ms = ConvertPDF.SetWatermark(stream, waterMark);

            }
            else if (ExcelExtensions.Any(t => t.Equals(fileExtension, StringComparison.OrdinalIgnoreCase)))
            {//excel 
                ms = ExcelWatermark.SetWatermark(stream, waterMark);
            }
            else
            {//其他文件
                byte[] tempBytes = new byte[stream.Length];
                stream.Read(tempBytes, 0, tempBytes.Length);
                ms.Write(tempBytes, 0, tempBytes.Length);
                ms.Close();
            }

            return ms;
        }
    }
}