using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace HtmlToOpenXml.Primitives
{   
    /// <summary>
    /// hierarchy node 
    /// </summary>
    internal class HieNode

    {
        int start=-1;
        int end=-1;
        string tag="";


  
        public HieNode(int start,int end,string tag)
        {
            this.start = start;
            this.end = end;
            this.tag = tag;
        }
    }
}
