using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    public class NumberingClass
    {
        private static NumberingClass instance = null;

        private NumberingClass()
        {
        }

        public static NumberingClass Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new NumberingClass();
                }
                return instance;
            }
        }
        public void FormatNumbering()
        {
            //Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
            //appWord = Interaction.GetObject(, "Word.Application");

            Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
            appWord = Globals.ThisAddIn.Application;
            // _Document doc = appWord.ActiveDocument;


            //foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in appWord.ActiveDocument.Paragraphs)
            //{
            //    Microsoft.Office.Interop.Word.Style _with1 = paragraph.get_Style() as Microsoft.Office.Interop.Word.Style;
            //    string styleName = _with1.NameLocal;
            //    string text = paragraph.Range.Text;
            //if (styleName == "Heading 1")
            //{
            var _with1 = appWord.ActiveDocument.Styles["Heading 1"];

            //var _with1 = appWord.ActiveDocument.Styles("Heading 1");  doubt
            //_with1.ParagraphFormat.Alignment = wdAlignParagraphJustify;
            //_with1.AutomaticallyUpdate = false;             doubt
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            appWord.ActiveDocument.UpdateStylesOnOpen = false;

                    //.BaseStyle = "Normal"
                    //_with1.NextParagraphStyle = "Heading 1";      doubt
                    // appWord.ActiveDocument.Styles.Add("Heading 1");

                    /*
                     _with1.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with1.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with1.ParagraphFormat.Alignment = wdAlignParagraphJustify;
                    _with1.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                     */
                    object setStyle = "Heading 1";
                    _with1.set_NextParagraphStyle(ref setStyle);
                    
            _with1.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
            _with1.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
            _with1.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            _with1.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);

           


            //.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            //.ParagraphFormat.LineSpacing = 24
            //.ParagraphFormat.SpaceAfter = 0
            //.ParagraphFormat.KeepWithNext = False
            _with1.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            _with1.ParagraphFormat.TabStops.ClearAll();
            _with1.Font.AllCaps = 0;
            _with1.Font.Bold = 0;
            _with1.Font.Italic = 0;
            _with1.Font.Underline = WdUnderline.wdUnderlineNone;

            // }
            //else if (styleName == "Heading 2")
            // {
            var _with2 = appWord.ActiveDocument.Styles["Heading 2"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle2 = "Heading 2";
            _with2.set_NextParagraphStyle(ref setStyle2);
            _with2.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
            _with2.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
            _with2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            _with2.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
            _with2.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            _with2.ParagraphFormat.TabStops.ClearAll();
            _with2.Font.AllCaps = 0;
            _with2.Font.Bold = 0;
            _with2.Font.Italic = 0;
            _with2.Font.Underline = WdUnderline.wdUnderlineNone;

            //}
            // else if (styleName == "Heading 3")
            //{
            var _with3 = appWord.ActiveDocument.Styles["Heading 3"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle3 = "Heading 3";
            _with3.set_NextParagraphStyle(ref setStyle3);
                    _with3.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with3.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with3.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with3.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    _with3.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
                    _with3.ParagraphFormat.TabStops.ClearAll();
                    _with3.Font.AllCaps = 0;
                    _with3.Font.Bold = 0;
                    _with3.Font.Italic = 0;
                    _with3.Font.Underline = WdUnderline.wdUnderlineNone;

            // }
            // else if (styleName == "Heading 4")
            //{
            var _with4 = appWord.ActiveDocument.Styles["Heading 4"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle4 = "Heading 4";
                    _with4.set_NextParagraphStyle(ref setStyle4);
                    _with4.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with4.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with4.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with4.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    _with4.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
                    _with4.ParagraphFormat.TabStops.ClearAll();
                    _with4.Font.AllCaps = 0;
                    _with4.Font.Bold = 0;
                    _with4.Font.Italic = 0;
                    _with4.Font.Underline = WdUnderline.wdUnderlineNone;

            //}
            //else if (styleName == "Heading 5")
            //{
            var _with5 = appWord.ActiveDocument.Styles["Heading 5"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle5 = "Heading 5";
                    _with5.set_NextParagraphStyle(ref setStyle5);
                    _with5.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with5.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with5.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with5.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    _with5.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel5;
                    _with5.ParagraphFormat.TabStops.ClearAll();
                    _with5.Font.AllCaps = 0;
                    _with5.Font.Bold = 0;
                    _with5.Font.Italic = 0;
                    _with5.Font.Underline = WdUnderline.wdUnderlineNone;
            //    }
            //  else if (styleName == "Heading 6")
            //{
            var _with6 = appWord.ActiveDocument.Styles["Heading 6"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle6 = "Heading 6";
                    _with6.set_NextParagraphStyle(ref setStyle6);
                    _with6.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with6.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with6.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with6.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    _with6.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel6;
                    _with6.ParagraphFormat.TabStops.ClearAll();
                    _with6.Font.AllCaps = 0;
                    _with6.Font.Bold = 0;
                    _with6.Font.Italic = 0;
                    _with6.Font.Underline = WdUnderline.wdUnderlineNone;

            //}
            //    else if (styleName == "Heading 7")
            //  {
            var _with7 = appWord.ActiveDocument.Styles["Heading 7"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle7 = "Body Text";
                    _with7.set_NextParagraphStyle(ref setStyle);
                    _with7.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with7.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with7.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with7.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    _with7.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel7;
                    _with7.ParagraphFormat.TabStops.ClearAll();
                    _with7.Font.AllCaps = 0;
                    _with7.Font.Bold = 0;
                    _with7.Font.Italic = 0;
                    _with7.Font.Underline = WdUnderline.wdUnderlineNone;
            //                }
            //      else if (styleName == "Heading 8")
            //    {
            var _with8 = appWord.ActiveDocument.Styles["Heading 8"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle8 = "Body Text";
                    _with8.set_NextParagraphStyle(ref setStyle8);
                    _with8.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with8.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with8.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with8.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    //_with1.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel7;
                    _with8.ParagraphFormat.TabStops.ClearAll();
                    _with8.Font.AllCaps = 0;
                    _with8.Font.Bold = 0;
                    _with8.Font.Italic = 0;
                    _with8.Font.Underline = WdUnderline.wdUnderlineNone;

            //}
            //  else if (styleName == "Heading 9")
            //{
            var _with9 = appWord.ActiveDocument.Styles["Heading 9"];
            appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    appWord.ActiveDocument.UpdateStylesOnOpen = false;
                    object setStyle9 = "Body Text";
                    _with9.set_NextParagraphStyle(ref setStyle9);
                    _with9.ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                    _with9.ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                    _with9.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    _with9.ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                    //_wih1.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel7;
                    _with9.ParagraphFormat.TabStops.ClearAll();
                    _with9.Font.AllCaps = 0;
                    _with9.Font.Bold = 0;
                    _with9.Font.Italic = 0;
                    _with9.Font.Underline = WdUnderline.wdUnderlineNone;

                //}

               
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].Reset(6);
                
                //appWord.ListGalleries(wdOutlineNumberGallery).Reset 7
                var _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1];
                _withList.NumberFormat = "%1.";
                _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                _withList.NumberPosition = appWord.InchesToPoints(1);
                _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                _withList.TextPosition = appWord.InchesToPoints(0);
                _withList.TabPosition = appWord.InchesToPoints(1.5F);
                _withList.ResetOnHigher = 0;
                _withList.StartAt = 1;
                // var _with2 = _with1.Font;
                _withList.Font.Bold = 0;
                _withList.Font.Italic = 0;
                _withList.Font.AllCaps = 0;
                _withList.Font.Size =(float)WdConstants.wdUndefined;
                _withList.Font.Animation =(WdAnimation) WdConstants.wdUndefined;
                _withList.Font.DoubleStrikeThrough =(int) WdUnderline.wdUnderlineNone;
                _withList.Font.Name = "";
                _withList.Font.Underline = WdUnderline.wdUnderlineNone;
                _withList.LinkedStyle = "Heading 1";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2];
            _withList.NumberFormat = "%2.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(1);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(1.5F);
            _withList.ResetOnHigher = 1;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = 0;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            //_withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 2";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3];
            _withList.NumberFormat = "%3.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(1.5f);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(2.0F);
            _withList.ResetOnHigher = 2;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = 0;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 3";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4];
            _withList.NumberFormat = "%4.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(2);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(2.5F);
            _withList.ResetOnHigher = 3;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = 0;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 4";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5];
            _withList.NumberFormat = "%5.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(2);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(2.5F);
            _withList.ResetOnHigher = 4;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = 0;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 5";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6];
            _withList.NumberFormat = "%6.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(2.5f);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(3);
            _withList.ResetOnHigher = 5;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = 0;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 6";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7];
            _withList.NumberFormat = "%7.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(3);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(1);
            _withList.TabPosition = appWord.InchesToPoints(3.5f);
            _withList.ResetOnHigher = 6;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps =(int) WdConstants.wdUndefined;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 7";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8];
            _withList.NumberFormat = "%8.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;

            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleNone;
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(0);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(0);
            _withList.ResetOnHigher = 7;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = (int)WdConstants.wdUndefined;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 8";

            _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9];
            _withList.NumberFormat = "%9.";
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;

            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleLowercaseRoman;
            _withList.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            _withList.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
            _withList.NumberPosition = appWord.InchesToPoints(0);
            _withList.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            _withList.TextPosition = appWord.InchesToPoints(0);
            _withList.TabPosition = appWord.InchesToPoints(0);
            _withList.ResetOnHigher = 8;
            _withList.StartAt = 1;
            // var _with2 = _with1.Font;
            _withList.Font.Bold = 0;
            _withList.Font.Italic = 0;
            _withList.Font.AllCaps = (int)WdConstants.wdUndefined;
            _withList.Font.Underline = (WdUnderline)WdConstants.wdUndefined;
            _withList.Font.Size = (float)WdConstants.wdUndefined;
            _withList.Font.Animation = (WdAnimation)WdConstants.wdUndefined;
            _withList.Font.DoubleStrikeThrough = (int)WdUnderline.wdUnderlineNone;
            _withList.Font.Name = "";
            _withList.Font.Underline = WdUnderline.wdUnderlineNone;
            _withList.LinkedStyle = "Heading 9";
            // }
        }
    }
}
