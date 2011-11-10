using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.util.collections;
using Org.BouncyCastle.Bcpg;

namespace test
{
   
    class MyHtml
    {
        static MyHtmlElement root = new MyHtmlElement();
        static MyHtmlElement dummy = new MyHtmlElement();

        private class MyHtmlElement
        {
            public IDictionary<string,string> Attributes { get; set; }
            public List<MyHtmlElement> Child { get; set; }
            public MyHtmlElement Parent { get; set; }
            public string Name { get; set; }
            public List<string> Comments { get; set; }
            public string Text { get; set; }
            public bool HasChilds { get; set; }
            public MyHtmlElement()
            {
                Attributes = new Dictionary<string, string>();
                Comments = new List<string>();
                Child = new List<MyHtmlElement>();
                HasChilds = true;
            }

            public override string ToString()
            {
                string s = "<" + Name;
                Attributes.ToList().ForEach(a=>s+=" "+a.Key+"=\""+a.Value+"\"");
                s+= ">" + Text;
                Child.ForEach(c=>s+=c+"\n");
                s += "</" + Name + ">";
                return s;
            }

            private string ExtractElementHeader(ref string html)
            {
                string header = html.Substring(0, html.IndexOf('>'));
                html = html.Remove(0, html.IndexOf('>'));
                return header;
            }
            private string ExtractElementText(ref string html)
            {
                string text = html.Substring(0, html.IndexOf('<')-1);
                html = html.Remove(0, html.IndexOf('>')-1);
                return text;
            }
            private void ExtractNameAndAttributes(string header)
            {
                var words = new List<string>(header.Split("\"".ToCharArray()));

                string sss = null;
                for (int i = 0; i < words.Count; i++)
                {
                    if (i % 2 == 1)
                    {
                        words[i] = words[i].Replace(' ', '§');
                        words[i] = words[i].Replace('=', '£');
                    }
                        

                    sss += words[i];
                }

                words = new List<string>(sss.Split(" =".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));

                Name = words[0];
                words.RemoveAt(0);

                if (words.Count % 2 == 1)
                    throw new Exception();

                for (int i = 1; i < words.Count; i += 2)
                {
                    Attributes.Add(words[i - 1], words[i].Replace("§", " ").Replace("£","="));
                }
            }
            private void ParseElementHeader(ref string html)
            {
                
                string header = html.Substring(1, html.IndexOf('>') - 1);
                html = html.Remove(0, html.IndexOf('>')+1);
             
                if (header.EndsWith("/")||header.StartsWith("meta")||header.StartsWith("br"))
                {
                    HasChilds = false;
                    header = header.TrimEnd(" /".ToCharArray());
                }

                ExtractNameAndAttributes(header);

                Text = html.Substring(0, html.IndexOf('<'));
                html = html.Substring(html.IndexOf('<'));
                
                if (!HasChilds)
                    return;

                if (html.StartsWith("</" + Name + ">"))
                {
                    HasChilds = false;
                    html = html.Substring(html.IndexOf('>')+1).TrimStart();
                }
            }
            
            private void ParseComments(ref string html)
            {
                string commentStart = "<!--";
                string commentEnd = "-->";
                while (html.Contains(commentStart))
                {
                    int start = html.IndexOf("<!--");
                    int end = html.IndexOf(commentEnd);
                    Comments.Add(html.Substring(start+commentStart.Length,end-start-commentStart.Length));
                    html = html.Remove(start, end-start+commentEnd.Length);
                }
            }

           
            public MyHtmlElement(ref string html, MyHtmlElement parent):this()
            {
                html = html.Replace("'", "\"");
                html = html.Replace("\r", "\n");
                html = html.Replace("\n", " ");
                Parent = parent;
                html = html.TrimStart();
               
                if (!html.StartsWith("<"))
                {
                    throw new Exception("Texto Html tem de começar por '<'");
                }

                if (html.StartsWith("</"))
                    return ;

                ParseComments(ref html);
                ParseElementHeader(ref html);


               if(HasChilds)
                {
                    string elementFooter = "</" + Name + '>';
                    while (!html.TrimStart().StartsWith(elementFooter))
                    {
                        var elem = new MyHtmlElement(ref html, this); 
                        Child.Add(elem);
                    }
                    html = html.Remove(0, elementFooter.Length);
                }

                //while (!html.Equals(""))
                //{
                //    Child.Add(new MyHtmlElement(ref html, Parent));
                //}
                //Console.WriteLine(html.Substring(0,html.IndexOf('\n')));

                html = html.TrimStart();
            }
        }

        /*
         * <html>
         *  <header>
         *  </header>
         *  <body>
         *  </body>
         * </html>
         * 
         */
        private static void HtmlDiet(string originFile, string targetFile)
        {
            StreamReader reader = new StreamReader(originFile,Encoding.GetEncoding(new CultureInfo("EN-US").TextInfo.ANSICodePage));
            StreamWriter writer = new StreamWriter(targetFile,false,Encoding.UTF8);

            string file = reader.ReadToEnd();

            MyHtmlElement page = new MyHtmlElement(ref file, root);
            
            writer.Write(page);
            writer.Flush();
            writer.Close();reader.Close();
        }

        public static void DietXmlDoc(string originFile, string targetFile = null)
        {
            string currentDir, dir, file;
            char dirChar = originFile.Contains('/') ? '/' : '\\';
            targetFile = targetFile ?? originFile;

            currentDir = Directory.GetCurrentDirectory();
            dir = originFile.Remove(originFile.LastIndexOf(dirChar));
            file = originFile.Remove(0, originFile.LastIndexOf(dirChar) + 1);

            Directory.SetCurrentDirectory(dir);

            XElement page = XElement.Load(file, LoadOptions.None);


            page.Elements().ToList().ForEach(xel => Console.WriteLine(xel));

            Directory.SetCurrentDirectory(currentDir);
        }


        public static void test()
        {
            HtmlDiet(@"C:\temp\test2.htm", @"C:\temp\test3.htm");
        }
    }
}
