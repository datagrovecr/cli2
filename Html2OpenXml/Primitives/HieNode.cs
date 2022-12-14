using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace HtmlToOpenXml.Primitives
{   
    /// <summary>
    /// Hierarchy node 
    /// </summary>
    internal class HieNode

    {
        public int start=-1;
        public int end =-1;
        public string tag ="";
        public int parent; //tells the index position of the parent
        private List<HieNode> child = new List<HieNode>();

        public HieNode()
        {
        }

        public HieNode(int parent)
        {
           this.parent = parent;
        }

        public HieNode(int start,int end,string tag)
        {
            this.start = start;
            this.end = end;
            this.tag = tag;
        }
    }
}
