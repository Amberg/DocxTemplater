﻿using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    public interface IFormatter
    {
        public bool CanHandle(Type type, string prefix);

        void ApplyFormat(string modelPath, object value, string prefix, string[] args, Text target);
    }
}