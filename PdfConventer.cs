﻿using System.IO;
using DinkToPdf;
using DinkToPdf.Contracts;

namespace PDFile
{
    public abstract class PdfConventer
    {
        public abstract byte[] ToPdf(MemoryStream ds);
        public abstract byte[] ToHtml(MemoryStream ds);

        protected static readonly IConverter Converter
            = new SynchronizedConverter(new PdfTools());
    }
}