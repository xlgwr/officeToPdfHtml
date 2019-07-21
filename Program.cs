using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using PDFile;

namespace PDF_Converter_demo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string filepath = AppDomain.CurrentDomain.BaseDirectory + @"Template\Template311.xls";
            MessageFile messageFile = new MessageFile()
            {
                FilePath = filepath,
                FileName = "Template311.xls",
                Message = "测试"
            };

            await ProcessMessagePdf(messageFile);
        }

        private static async Task<string> ProcessMessagePdf(MessageFile messageFile)
        {
            PdfConventer _pdfConventer;
            MemoryStream _downloadStream = new MemoryStream();
            var message = messageFile.Message;
            if (messageFile.FilePath == null)
                throw new Exception("Send the document with the extension .docx or .xls or .xlsx ");
            var filePath = messageFile.FilePath;
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(messageFile.FileName);
            var extension = Path.GetExtension(messageFile.FileName);
            _downloadStream = await ReadFileToMemoryStream(filePath);
            switch (extension)
            {
                case ".xls":
                case ".xlsx":

                    _pdfConventer = new XlsxToPdf();
                    Console.WriteLine(_pdfConventer.GetType().ToString());
                    break;
                case ".docx":
                    _pdfConventer = new DocxToPdf();
                    Console.WriteLine(_pdfConventer.GetType().ToString());
                    break;
                default:
                    throw new Exception("There is no such format. Supported Formats:docx,xls,xlsx ");
            }

            // return await SendFilePdf(fileNameWithoutExtension, _pdfConventer, _downloadStream);
            return await SendFileHtml(fileNameWithoutExtension, _pdfConventer, _downloadStream);
        }
        static async Task<string> SendFileHtml(string fileNameWithoutExtension, PdfConventer pdf, MemoryStream ds)
        {
            Console.WriteLine("...");

            Console.WriteLine(fileNameWithoutExtension);
            var filePath = fileNameWithoutExtension + ".html";
            var stream = new MemoryStream(pdf.ToHtml(ds));

            await saveTofle(stream, filePath);

            Console.WriteLine("!!!" + filePath);

            return filePath;
        }
        static async Task<string> SendFilePdf(string fileNameWithoutExtension, PdfConventer pdf, MemoryStream ds)
        {
            Console.WriteLine("...");

            Console.WriteLine(fileNameWithoutExtension);
            var filePath = fileNameWithoutExtension + ".pdf";
            var stream = new MemoryStream(pdf.ToPdf(ds));

            await saveTofle(stream, filePath);

            Console.WriteLine("!!!" + filePath);

            return filePath;
        }
        public static async Task saveTofle(MemoryStream file, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write))
            {
                byte[] buffer = file.ToArray();//转化为byte格式存储
                await fs.WriteAsync(buffer, 0, buffer.Length);
                fs.Flush();
                buffer = null;
            }//使用using可以最后不用关闭fs 比较方便
        }
        //流Stream 和 文件file之间的转换

        //将 Stream 写入文件
        public static async Task stream2file(Stream stream, string fileName)
        {

            // 把 Stream 转换成 byte[]
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                // 把 byte[] 写入文件
                await fs.WriteAsync(bytes, 0, bytes.Length);
                // BinaryWriter bw = new BinaryWriter(fs);
                // bw.Write(bytes);
                // bw.Close();
                //fs.Close();  
            }

        }


        //从文件读取 Stream
        public static async Task<Stream> file2stream(string fileName)
        {

            // 打开文件
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // 读取文件的 byte[]
                byte[] bytes = new byte[fileStream.Length];
                await fileStream.ReadAsync(bytes, 0, bytes.Length);
                // 把 byte[] 转换成 Stream
                Stream stream = new MemoryStream(bytes);
                return stream;
            }


        }

        //字节数组到字符串的解码
        public static string str2byte(byte[] data)
        {
            string str = System.Text.Encoding.UTF8.GetString(data);
            //str = Convert.ToBase64String(data);
            //有很多种编码方式，可参考:http://blog.csdn.net/luanpeng825485697/article/details/77622243
            return str;
        }

        /****字节数组byte[]与内存流MemoryStream之间的转换****/
        //字节数组转化为输入内存流
        public static MemoryStream byte2stream(byte[] data)
        {
            MemoryStream inputStream = new MemoryStream(data);
            return inputStream;
        }
        public static async Task<MemoryStream> ReadFileToMemoryStream(string filePath)
        {
            byte[] data = await File.ReadAllBytesAsync(filePath);
            MemoryStream ms = new MemoryStream(data);
            return ms;
        }


        //输出内存流转化为字节数组
        public static byte[] byte2stream(MemoryStream outStream)
        {
            return outStream.ToArray();
        }

        /************字节数组byte[]与流stream之间的转换********/
        //将 Stream 转成 byte[]
        public byte[] stream2byte(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }
    }
}




