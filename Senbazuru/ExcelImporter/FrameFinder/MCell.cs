using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace FrameFinder
{
    class MCell
    {
        private static char[] punc = { '!', '\"', '#', '$', '%', '&', '\'', '(', ')',
                                         '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', '>', '?',
                                         '@', '[', '\\', ']', '^', '_', '`', '{', '|', '}', '~' };
        private static HashSet<char> punctuation = new HashSet<char>(punc);
        private string value;
        private string type;
        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
        public string Type
        {
            get { return this.type; }
            set { this.type = value; }
        }

        private bool rightAlignFlag;
        private bool leftAlignFlag;
        private bool centerAlignFlag;
        public bool RightAlignFlag
        {
            get { return this.rightAlignFlag; }
        }
        public bool LeftAlignFlag
        {
            get { return this.leftAlignFlag; }
        }
        public bool CenterAlignFlag
        {
            get { return this.centerAlignFlag; }
        }


        private bool bottomBorder;
        private bool upperBorder;
        private bool leftBorder;
        private bool rightBorder;
        public bool BottomBorder
        {
            get { return this.bottomBorder; }
        }
        public bool UpperBorder
        {
            get { return this.upperBorder; }
        }
        public bool LeftBorder
        {
            get { return this.leftBorder; }
        }
        public bool RightBorder
        {
            get { return this.rightBorder; }
        }
        
        private int bgColor;
        private int height;
        private int italicFlag;
        private int underlineFlag;
        private bool boldFlag;
        public int BackgroundColor
        {
            get { return this.bgColor; }
        }
        public int FontHeight
        {
            get { return this.height; }
        }
        public int ItalicFlag
        {
            get { return this.italicFlag; }
        }
        public int UnderlineFlag
        {
            get { return this.underlineFlag; }
        }
        public bool BoldFlag
        {
            get { return this.boldFlag; }
        }

        private int indents;
        public int Indents
        {
            get { return this.indents; }
        }
        
        public MCell() { }
        public void Init(string value, string cType, int indents, int alignStyle,
            int boldFlag, string borderStyle, int bgColor, int height,
            int italicFlag, int underlineFlag)
        {
            this.value = value;
            this.type = cType;
            this.indents = this.getIndents(indents);
            this.centerAlignFlag = false;
            this.leftAlignFlag = false;
            this.rightAlignFlag = false;

            /* XlHAlign
             * -4131 = Left   -> ALIGN_LEFT = 0x1
             * -4152 = Right  -> ALIGN_RIGHT = 0x3
             * -4108 = Center -> ALIGN_CENTER = 0x2
             */
            if (alignStyle == (int)XlHAlign.xlHAlignLeft)
            {
                this.leftAlignFlag = true;
            }
            else if (alignStyle == (int)XlHAlign.xlHAlignCenter)
            {
                this.centerAlignFlag = true;
            }
            else if (alignStyle == (int)XlHAlign.xlHAlignRight)
            {
                this.rightAlignFlag = true;
            }

            this.boldFlag = (boldFlag == 1 ? true : false);
            this.bottomBorder = (borderStyle[0] == '1' ? true : false);
            this.upperBorder = (borderStyle[1] == '1' ? true : false);
            this.leftBorder = (borderStyle[2] == '1' ? true : false);
            this.rightBorder = (borderStyle[3] == '1' ? true : false);
            this.bgColor = bgColor;
            this.height = height;
            this.italicFlag = italicFlag;
            this.underlineFlag = underlineFlag;
        }

        private int getIndents(int indents)
        {
            if (this.value.Length == 0)
            {
                return 0;
            }
            int i = 0;
            for (; i < this.value.Length; ++i)
            {
                if (this.value[i] == ' ' || punctuation.Contains(this.value[i]))
                {
                    continue;
                }
                else
                {
                    break;
                }
            }
            return i + indents * 2;
        }

        public char WriteStrAlignStyle()
        {
            if (this.leftAlignFlag)
            {
                return '1';
            }
            else if (this.centerAlignFlag)
            {
                return '2';
            }
            else if (this.rightAlignFlag)
            {
                return '3';
            }
            else
            {
                return '0';
            }
        }

        public String WriteStrBordStyle()
        {
            String ret = "";
            ret += (this.bottomBorder ? "1" : "0");
            ret += (this.upperBorder ? "1" : "0");
            ret += (this.leftBorder ? "1" : "0");
            ret += (this.rightBorder ? "1" : "0");
            return ret;
        }

        public void PrintInfo()
        {
            Console.WriteLine(this.value + " " + this.type + " " + this.indents);
            Console.WriteLine("bold: " + this.boldFlag);
            Console.WriteLine("align: " + this.leftAlignFlag + " " + this.centerAlignFlag + " " + this.rightAlignFlag);
            Console.WriteLine("border: " + this.bottomBorder + " " + this.upperBorder + " " + this.leftBorder + " " + this.rightBorder);
        }

        public override string ToString()
        {
            string ret = "";
            ret += this.value + " " + this.type + " " + this.indents + "\n";
            ret += "bold: " + this.boldFlag + "\n";
            ret += "align: " + this.leftAlignFlag + " " + this.centerAlignFlag + " " + this.rightAlignFlag + "\n";
            ret += "border: " + this.bottomBorder + " " + this.upperBorder + " " + this.leftBorder + " " + this.rightBorder + "\n";
            return ret;
        }
    }
}
