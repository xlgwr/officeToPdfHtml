﻿using System;
using System.IO;
using System.Text;
using System.Xml;
using DinkToPdf;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;

namespace PDFile
{
    public class XlsxToPdf : PdfConventer
    {
        public override byte[] ToHtml(MemoryStream downloadStream)
        {

            downloadStream.Position = 0;
            var excelToHtmlConverter = new ExcelToHtmlConverter
            {
                OutputColumnHeaders = false,
                OutputHiddenColumns = false,
                OutputHiddenRows = false,
                OutputLeadingSpacesAsNonBreaking = false,
                OutputRowNumbers = false,
                UseDivsToSpan = false
            };
            excelToHtmlConverter.ProcessWorkbook(WorkbookFactory.Create(downloadStream));
            Console.WriteLine("Coverted MemoryStream xlsx в html");
            
            MemoryStream msXml = new MemoryStream();
            XmlWriter xmlWriter = new XmlTextWriter(msXml, Encoding.UTF8);
            excelToHtmlConverter.Document.WriteContentTo(xmlWriter);

            return msXml.ToArray();
        }
        public override byte[] ToPdf(MemoryStream downloadStream)
        {
            downloadStream.Position = 0;
            var excelToHtmlConverter = new ExcelToHtmlConverter
            {
                OutputColumnHeaders = false,
                OutputHiddenColumns = false,
                OutputHiddenRows = false,
                OutputLeadingSpacesAsNonBreaking = false,
                OutputRowNumbers = false,
                UseDivsToSpan = false
            };
            excelToHtmlConverter.ProcessWorkbook(WorkbookFactory.Create(downloadStream));
            Console.WriteLine("Coverted MemoryStream xlsx в html");
            var doc = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    Orientation = Orientation.Landscape
                },
                Objects =
                {
                    new ObjectSettings
                    {
                        HtmlContent = excelToHtmlConverter.Document.OuterXml,
                        WebSettings = {DefaultEncoding = "utf-8"}
                    }
                }
            };
            Console.WriteLine("Clear html");
            var bytes = Converter.Convert(doc);
            Console.WriteLine("Convert html to pdf");
            return bytes;
        }
    }
}