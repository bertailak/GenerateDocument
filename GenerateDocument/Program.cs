using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            string utilPath = Directory.GetCurrentDirectory() + "/";
            string sourcePath = utilPath + "GenerateDocumentTemp/";
            GenerateDocument gd = new GenerateDocument(sourcePath, utilPath);

            try
            {
                string command = "";
                string filename = "";

                #region Get command and filename

                string xmlPath = sourcePath + "EmptyXML.xml";

                if (args.Length > 0)
                {
                    command = args[0];
                }
                else
                {
                    Console.Write("Command: ");
                    command = Console.ReadLine();
                }

                if (!command.Equals("help"))
                {
                    if (args.Length > 1)
                    {
                        filename = args[1];
                    }
                    else
                    {
                        Console.Write("Filename: ");
                        filename = Console.ReadLine();
                    }
                }

                Console.WriteLine("");

                //command = "DOCXtoXSLT";
                //filename = "AgreePKB.docx";

                #endregion

                #region Select method by command
                switch (command)
                {
                    case "xslt":
                    case "xslt:s":
                    case "xslt:o":
                    case "xslt:s:o":
                    case "xslt:o:s":
                    case "docx":
                    case "docx:o":
                    case "docx:s":
                    case "docx:s:o":
                    case "docx:o:s":
                    case "pdf":
                    case "pdf:o":
                    case "pdf:s":
                    case "pdf:s:o":
                        gd.RunCommand(command, filename);
                        break;
                    case "xml":
                        gd.GenerateXMLfromXSLT(filename);
                        break;
                    case "tdocx":
                    case "tdocx:o":
                        string result = gd.GenerateDOCXfromXSLT(filename);
                        gd.OpenFile(command, result);
                        break;
                    case "help":
                        GetCommands();
                        break;
                    default:
                        GetDefaultCommand(command);
                        break;
                }
                #endregion

                Console.WriteLine("~~Success!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: \n" + ex.Message);
            }
            finally
            {
                //Console.WriteLine("~~Press any key!");
                //Console.ReadKey();
            }

        }

        static void GetCommands()
        {
            Console.WriteLine("~~Versions:");
            Console.WriteLine("- GenerateDocument:   1.3");
            Console.WriteLine("- xsl:                   1.0");
            Console.WriteLine();
            //Console.WriteLine("~~Please write available commands: ");
            Console.WriteLine("~~COMMAND        FileFormat");
            Console.WriteLine("----------------------------");
            Console.WriteLine("- xslt           *.docx");
            Console.WriteLine("- docx           *.docx");
            Console.WriteLine("- pdf            *.docx");
            Console.WriteLine("- docx           *.xslt");
            Console.WriteLine("- pdf            *.xslt");
            Console.WriteLine("- xml            *.xslt");
            Console.WriteLine("- tdocx          *.xslt");
            //Console.WriteLine("EmptyXML");
            Console.WriteLine();
            Console.WriteLine("~~Attribute      Description");
            Console.WriteLine("----------------------------");
            Console.WriteLine("- :o             open");
            //Console.WriteLine("- :s             xslt save");
            Console.WriteLine();
        }
        static void GetDefaultCommand(string command)
        {
            throw new Exception("'" + command + "' is not a GenerateDocument command. See 'help'!");
        }
    }
}
