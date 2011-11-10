using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using System.util.collections;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace test
{
    class MyOffice
    {
    
        public static IDictionary<string,string> ReadExcel(string file, int sheet = 1, int keyColumn = 1, int valueColumn = 2, int fromRow = 0, int toRow = int.MaxValue)
        {
            Dictionary<string,string> dict = new Dictionary<string, string>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);
            xlWorkbook.Activate();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            
            for (int row = 0; row < rowCount && row < toRow; row++)
            {
                try
                {
                    dict.Add(xlRange.Cells[row, keyColumn].Value2.ToString(), xlRange.Cells[row, valueColumn].Value2.ToString());
                }
                catch (Exception)
                {
                }
            }
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            return dict;

        }

        public static void SaveWordDoc(string originFile, IEnumerable<KeyValuePair<string, string>> replaceList = null, string targetFile = null, WdSaveFormat format = WdSaveFormat.wdFormatDocumentDefault, MsoEncoding encoding = MsoEncoding.msoEncodingAutoDetect)
        {
            Word.Application ap = null;
            Word.Document doc = null;
            object missing = Type.Missing;
            bool success = true;
            replaceList = replaceList ?? new Dictionary<string, string>();
            targetFile = targetFile ?? originFile;
            if(targetFile.LastIndexOf('.')>targetFile.LastIndexOf('/')||targetFile.LastIndexOf('.')>targetFile.LastIndexOf('\\'))
                targetFile=targetFile.Remove(targetFile.LastIndexOf('.'));
                
            try
            {
                ap = new Word.Application();
                ap.DisplayAlerts = WdAlertLevel.wdAlertsNone;       
                doc = ap.Documents.Open(originFile, ReadOnly: false, Visible: false);
                doc.Activate();
                
                Selection sel = ap.Selection;

                if (sel == null) 
                    throw new Exception("Unable to acquire Selection...no writing to document done..");
                
                switch (sel.Type)
                {
                    case WdSelectionType.wdSelectionIP:
                        replaceList.ToList().ForEach(p => sel.Find.Execute(FindText: p.Key, ReplaceWith: p.Value, Replace: WdReplace.wdReplaceAll));
                        break;
                    default:
                        throw new Exception("Selection type not handled; no writing done");
                }

                sel.Paragraphs.LineUnitAfter = 0;
                sel.Paragraphs.LineUnitBefore = 0;
                sel.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            
                doc.SaveSubsetFonts = false;
                doc.SaveAs(targetFile, format, Encoding: encoding);
               
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(doc);
                }
                if (ap != null)
                {
                    ap.Quit(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(ap);
                }
                if(!success)
                    throw new Exception(); // Could be that the document is already open (/) or Word is in Memory(?)   
            }
        }
        //    //if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\wordTemp"))
        //    //    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\wordTemp");

        //            tempfilename = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\wordTemp\\" + count + ".doc";
        //            wrdDoc.SaveAs(ref tempfilename, ref missing, ref missing, ref missing, ref missing,
        //                 ref missing, ref missing, ref missing, ref missing,
        //                  ref missing, ref missing, ref missing, ref missing,
        //                   ref missing, ref missing, ref missing);
        //            files.Add(tempfilename.ToString());
        //            count++;
        //            wrdDoc.Close(ref missing, ref missing, ref missing);
       
        public static void testExcel()
        {
            var dict = ReadExcel(@"C:/temp/test.xls");
            dict.ToList().ForEach(p=>Console.WriteLine(p.Key+"->"+p.Value));
            Console.WriteLine("Done!");
        }

        public static void testWord()
        {
            Dictionary<string, string> replaceDict = new Dictionary<string, string>();
            replaceDict.Add("<<loco>>", "LE5666");
            replaceDict.Add("<<datainicio>>", "11-11-11");
            replaceDict.Add("<<horainicio>>", "11:11");
            replaceDict.Add("<<datafim>>", "22-22-22");
            replaceDict.Add("<<horafim>>", "22:22");
            replaceDict.Add("<<operador>>", "Gaspar Francisco");
            replaceDict.Add("<<datacontacto>>", "33-33-33");
            replaceDict.Add("<<horacontacto>>", "33:33");
            replaceDict.Add("<<visita>>", "VS+VAV");
            replaceDict.Add("<<km>>", "999999");
            replaceDict.Add("<<horas>>", "88888");
            replaceDict.Add("<<kwt>>", "777777");
            replaceDict.Add("<<kwf>>", "666666");
            
            SaveWordDoc(@"C:\temp\test2.docx",format:WdSaveFormat.wdFormatFilteredHTML,replaceList:replaceDict);
            Console.WriteLine("Done!");
        }
    }
}
