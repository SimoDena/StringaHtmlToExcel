using Infragistics.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace stringaHTMLtoExcel
{
    public class Carattere
    {
        public char c;
        public bool underline;
        public bool bold;
        public bool italic;

        public Carattere(char _c, bool _underline, bool _bold, bool _italic)
        {
            c = _c;
            underline = _underline;
            bold = _bold;
            italic = _italic;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string sHTML = "A<b><u>BCD</u><i>EF</i>w</b>G";
            HtmlExcel(sHTML);
            Console.ReadKey();
        }
        //static List<Carattere> ScorriLista(List<Carattere> stringaNoHTML, string tmp, bool nowBold, bool nowUnder, bool nowItalic, int i)
        //{
        //    while (i < tmp.Length)
        //    {
        //        if (tmp[i] == '<' && tmp[i + 1] == 'b' && tmp[i + 2] == '>')
        //        {
        //            i = i + 3;
        //            nowBold = true;
        //        }
        //        if (tmp[i] == '<' && tmp[i + 1] == 'u' && tmp[i + 2] == '>')
        //        {
        //            i = i + 3;
        //            nowUnder = true;
        //        }
        //        if (tmp[i] == '<' && tmp[i + 1] == 'i' && tmp[i + 2] == '>')
        //        {
        //            i = i + 3;
        //            nowItalic = true;
        //        }
        //        if (tmp[i] == '<' && tmp[i + 1] == '/' && tmp[i + 2] == 'b' && tmp[i + 3] == '>')
        //        {
        //            i = i + 4;
        //            nowBold = false;
        //            ScorriLista(stringaNoHTML, tmp, nowBold, nowUnder, nowItalic, i);

        //        }
        //        if (tmp[i] == '<' && tmp[i + 1] == '/' && tmp[i + 2] == 'u' && tmp[i + 3] == '>')
        //        {
        //            i = i + 4;
        //            nowUnder = false;
        //            ScorriLista(stringaNoHTML, tmp, nowBold, nowUnder, nowItalic, i);
        //        }
        //        if (tmp[i] == '<' && tmp[i + 1] == '/' && tmp[i + 2] == 'i' && tmp[i + 3] == '>')
        //        {
        //            i = i + 4;
        //            nowItalic = false;
        //            ScorriLista(stringaNoHTML, tmp, nowBold, nowUnder, nowItalic, i);
        //        }

        //        Carattere car = new Carattere(tmp[i], nowUnder, nowBold, nowItalic);
        //        stringaNoHTML.Add(car);
        //        i++;
        //    }
        //    return stringaNoHTML;
        //}
        static void _HtmlExcel(string sHTML)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets.Add("ws1");

            List<Carattere> stringaNoHTML = new List<Carattere>();
            bool nowBold = false;
            bool nowUnder = false;
            bool nowItalic = false;

            //stringaNoHTML = ScorriLista(stringaNoHTML, tmp, nowBold, nowUnder, nowItalic,i);

            int i = 0;
            bool ripeti = false;
            while (i < sHTML.Length)
            {
                try
                {
                    ripeti = false;
                    if (sHTML[i] == '<' && sHTML[i + 1] == 'b' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowBold = true;
                    }
                    if (sHTML[i] == '<' && sHTML[i + 1] == 'u' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowUnder = true;
                        ripeti = true;
                    }
                    if (sHTML[i] == '<' && sHTML[i + 1] == 'i' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowItalic = true;
                        ripeti = true;
                    }
                    if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'b' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowBold = false;
                        ripeti = true;

                    }
                    if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'u' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowUnder = false;
                        ripeti = true;
                    }
                    if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'i' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowItalic = false;
                        ripeti = true;
                    }

                    if (ripeti == false)
                    {
                        Carattere car = new Carattere(sHTML[i], nowUnder, nowBold, nowItalic);
                        stringaNoHTML.Add(car);
                        i++;
                    }
                }
                catch
                {
                    break;
                }
            }

            char[] copiaCharHtml = new char[stringaNoHTML.Count()];
            for (i = 0; i < stringaNoHTML.Count(); i++)
            {
                copiaCharHtml[i] = stringaNoHTML[i].c;
            }
            string copiaStringaNoHTML = new string(copiaCharHtml);
            FormattedString fs = new FormattedString(copiaStringaNoHTML);
            wb.Worksheets[0].Rows[1].Cells[1].Value = fs;
            for (i = 0; i < stringaNoHTML.Count(); i++)
            {
                if (stringaNoHTML[i].bold == true)
                {
                    fs.GetFont(i, 1).Bold = ExcelDefaultableBoolean.True;
                }
                if (stringaNoHTML[i].italic == true)
                {
                    fs.GetFont(i, 1).Italic = ExcelDefaultableBoolean.True;
                }
                if (stringaNoHTML[i].underline == true)
                {
                    fs.GetFont(i, 1).UnderlineStyle = FontUnderlineStyle.Single;
                }
            }
            wb.Worksheets[0].Rows[1].Cells[1].Value = fs;
            wb.Save("fileExcel.xls");
        }

        static void HtmlExcel(string sHTML)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets.Add("ws1");

            List<Carattere> stringaNoHTML = new List<Carattere>();
            bool nowBold = false;
            bool nowUnder = false;
            bool nowItalic = false;

            //stringaNoHTML = ScorriLista(stringaNoHTML, tmp, nowBold, nowUnder, nowItalic,i);

            int i = 0;
            while (i < sHTML.Length)
            {
                try
                {
                    if (sHTML[i] == '<' && sHTML[i + 1] == 'b' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowBold = true;
                    }
                    else if (sHTML[i] == '<' && sHTML[i + 1] == 'u' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowUnder = true;
                    }
                    else if (sHTML[i] == '<' && sHTML[i + 1] == 'i' && sHTML[i + 2] == '>')
                    {
                        i = i + 3;
                        nowItalic = true;
                    }
                    else if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'b' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowBold = false;

                    }
                    else if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'u' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowUnder = false;
                    }
                    else if (sHTML[i] == '<' && sHTML[i + 1] == '/' && sHTML[i + 2] == 'i' && sHTML[i + 3] == '>')
                    {
                        i = i + 4;
                        nowItalic = false;
                    }
                    else
                    {
                        Carattere car = new Carattere(sHTML[i], nowUnder, nowBold, nowItalic);
                        stringaNoHTML.Add(car);
                        i++;
                    }
                }
                catch
                {
                    break;
                }
            }

            char[] copiaCharHtml = new char[stringaNoHTML.Count()];
            for (i = 0; i < stringaNoHTML.Count(); i++)
            {
                copiaCharHtml[i] = stringaNoHTML[i].c;
            }
            string copiaStringaNoHTML = new string(copiaCharHtml);
            FormattedString fs = new FormattedString(copiaStringaNoHTML);
            wb.Worksheets[0].Rows[1].Cells[1].Value = fs;
            for (i = 0; i < stringaNoHTML.Count(); i++)
            {
                if (stringaNoHTML[i].bold == true)
                {
                    fs.GetFont(i, 1).Bold = ExcelDefaultableBoolean.True;
                }
                if (stringaNoHTML[i].italic == true)
                {
                    fs.GetFont(i, 1).Italic = ExcelDefaultableBoolean.True;
                }
                if (stringaNoHTML[i].underline == true)
                {
                    fs.GetFont(i, 1).UnderlineStyle = FontUnderlineStyle.Single;
                }
            }
            wb.Worksheets[0].Rows[1].Cells[1].Value = fs;
            wb.Save("fileExcel.xls");
        }
    }
}
