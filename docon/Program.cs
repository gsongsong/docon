using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace docon
{
  class Program
  {
    static readonly string FormatTxt = "txt";
    static readonly string FormatHtm = "htm";
    static readonly int EncodingUtf8 = 65001;

    static void Main(string[] args)
    {
      CheckArgs(args);
      var inputFile = args[0];
      CheckInputFile(inputFile);
      var nameWoExt = Path.GetFileNameWithoutExtension(inputFile);
      var outputFormat = args[1];
      CheckOutputFormat(outputFormat);
      var currentDirectory = Directory.GetCurrentDirectory();
      var outputFile = Path.Combine(currentDirectory, nameWoExt + "." + outputFormat);
      var wordApp = new Word.Application();
      var doc = wordApp.Documents.Open(Path.Combine(currentDirectory, inputFile));
      ConvertArrowToUnicode(doc);
      var format = outputFormat == FormatTxt ? WdSaveFormat.wdFormatText : WdSaveFormat.wdFormatFilteredHTML;
      doc.SaveAs2(outputFile, format, Encoding: EncodingUtf8);
      doc.Close();
      wordApp.Quit();
    }

    static void CheckArgs(string[] args)
    {
      if (args.Length != 2)
      {
        Console.WriteLine("Input file and output format must be specified. Exit.");
        Environment.Exit(-1);
      }
    }

    static void CheckInputFile(string inputFile)
    {
      if (!File.Exists(inputFile))
      {
        Console.WriteLine("Input file does not exist. Exit.");
        Environment.Exit(-1);
      }
    }

    static void CheckOutputFormat(string outputFormat)
    {
      if (outputFormat != FormatTxt && outputFormat != FormatHtm)
      {
        Console.WriteLine("Output format must be either \"txt\" or \"htm\"");
        Environment.Exit(-1);
      }
    }

    static void ConvertArrowToUnicode(Document doc)
    {
      var from = ((char)0xF0AE).ToString();
      var to = ((char)0x2192).ToString();
      doc.Content.Find.Execute(from, ReplaceWith: to, Replace: WdReplace.wdReplaceAll);
    }
  }
}
