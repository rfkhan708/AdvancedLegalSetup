using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    /// <summary>
    /// 
    /// </summary>
    public class GlobalEnumClass
    {
        private static GlobalEnumClass instance = null;
        /// <summary>
        /// 
        /// </summary>
        private GlobalEnumClass()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        public static GlobalEnumClass Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new GlobalEnumClass();
                }
                return instance;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public enum AllEnum
        {
            False=0,
            True=1
        }
        public const string Heading1 = "Heading 1";
        public const string Heading2 = "Heading 2";
        public const string Heading3 = "Heading 3";
        public const string Heading4 = "Heading 4";
        public const string Heading5 = "Heading 5";
        public const string Heading6 = "Heading 6";
        public const string Heading7 = "Heading 7";
        public const string Heading8 = "Heading 8";
        public const string Heading9 = "Heading 9";
        public bool GlobalTemplate = false;
        public const string PortlandLetterheadforblankpaper = "Portland Letterhead for blank paper";
        public const string CentralOregonLetterheadforblankpaper = "Central Oregon Letterhead for blank paper";
    }
}
