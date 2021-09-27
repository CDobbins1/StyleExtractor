using System;
using System.IO;

namespace StyleExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter the file path of the .docx: ");
            string docPath = Console.ReadLine();
            if (File.Exists(docPath))
            {
                if (Path.GetExtension(docPath) == ".docx")
                {
                    StyleExtractor.PrintStyles(docPath);
                }
                else
                {
                    Console.WriteLine("File is not of type .docx!");
                }
            }
            else
            {
                Console.WriteLine("File not found!");
            }
        }
    }
}
