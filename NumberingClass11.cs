using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    public class NumberingClass11
    {
        private static NumberingClass11 instance = null;
        private NumberingClass11()
        {
        }
        public static NumberingClass11 Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new NumberingClass11();
                }
                return instance;
            }
        }
        /// <summary>
        /// For Calling FormatNumbering method in FormatNumbering content under legal Ribbon
        /// </summary>
        public void FormatNumbering()
        {
            StringBuilder sbTrace = new StringBuilder();
            try
            {

                sbTrace.AppendLine("Start");
                Logger.SaveLoggerTrace(sbTrace);
                Microsoft.Office.Interop.Word.Application appWord = default(Microsoft.Office.Interop.Word.Application);
                appWord = Globals.ThisAddIn.Application;
                appWord.ActiveDocument.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.UpdateStylesOnOpen = false;
                //.BaseStyle = "Normal"
                object setStyle = GlobalEnumClass.Heading1;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].Font.Bold = (int)GlobalEnumClass.AllEnum.True;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading1].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ActiveDocument.UpdateStylesOnOpen = false;

                setStyle = GlobalEnumClass.Heading2;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading2].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ActiveDocument.UpdateStylesOnOpen = false;

                setStyle = GlobalEnumClass.Heading3;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading3].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ActiveDocument.UpdateStylesOnOpen = false;

                setStyle = GlobalEnumClass.Heading4;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading4].Font.Underline = WdUnderline.wdUnderlineNone;

                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].AutomaticallyUpdate = false;
                //.BaseStyle = "Normal"
                setStyle = GlobalEnumClass.Heading5;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel5;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading5].Font.Underline = WdUnderline.wdUnderlineNone;

                // _with1 = appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6];
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].AutomaticallyUpdate = false;
                //.BaseStyle = "Normal"
                setStyle = GlobalEnumClass.Heading6;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel6;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading6].Font.Underline = WdUnderline.wdUnderlineNone;

                // _with1 = appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7];
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].AutomaticallyUpdate = false;
                //.BaseStyle = "Normal"
                setStyle = "Body Text";
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel7;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading7].Font.Underline = WdUnderline.wdUnderlineNone;

                //_with1 = appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8];
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].AutomaticallyUpdate = false;
                //.BaseStyle = "Normal"
                setStyle = "Body Text";
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel8;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading8].Font.Underline = WdUnderline.wdUnderlineNone;

                //_with1 = appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9];
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].AutomaticallyUpdate = false;
                //.BaseStyle = "Normal"
                setStyle = "Body Text";
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].set_NextParagraphStyle(ref setStyle);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.LeftIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.RightIndent = appWord.InchesToPoints(0);
                //appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.FirstLineIndent = appWord.InchesToPoints(0);
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel9;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].ParagraphFormat.TabStops.ClearAll();
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ActiveDocument.Styles[GlobalEnumClass.Heading9].Font.Underline = WdUnderline.wdUnderlineNone;

                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].Reset(6);

                //var _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].NumberFormat = "%1.";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].NumberPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].TabPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].ResetOnHigher = 0;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Bold = (int)GlobalEnumClass.AllEnum.True;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[1].LinkedStyle = GlobalEnumClass.Heading1;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].NumberFormat = "(%2)";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].NumberPosition = appWord.InchesToPoints(1);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].TabPosition = appWord.InchesToPoints(1.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].ResetOnHigher = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[2].LinkedStyle = GlobalEnumClass.Heading2;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].NumberFormat = "(%3)";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].NumberStyle = WdListNumberStyle.wdListNumberStyleLowercaseLetter;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].NumberPosition = appWord.InchesToPoints(1);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].TabPosition = appWord.InchesToPoints(1.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].ResetOnHigher = 2;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[3].LinkedStyle = GlobalEnumClass.Heading3;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].NumberFormat = "%4";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].NumberPosition = appWord.InchesToPoints(1);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].TabPosition = appWord.InchesToPoints(1.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].ResetOnHigher = 3;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Animation =(WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[4].LinkedStyle = GlobalEnumClass.Heading4;

                // _withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].NumberFormat = "(%5)";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].NumberPosition = appWord.InchesToPoints(2);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].TextPosition = appWord.InchesToPoints(0.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].TabPosition = appWord.InchesToPoints(2.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].ResetOnHigher = 4;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[5].LinkedStyle = GlobalEnumClass.Heading5;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].NumberFormat = "%6";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].NumberStyle = WdListNumberStyle.wdListNumberStyleLowercaseLetter;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].NumberPosition = appWord.InchesToPoints(2.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].TextPosition = appWord.InchesToPoints(1);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].TabPosition = appWord.InchesToPoints(3);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].ResetOnHigher = 5;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[6].LinkedStyle = GlobalEnumClass.Heading6;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].NumberFormat = "%7";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].NumberPosition = appWord.InchesToPoints(3);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].TextPosition = appWord.InchesToPoints(1);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].TabPosition = appWord.InchesToPoints(3.5f);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].ResetOnHigher = 6;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[7].LinkedStyle = GlobalEnumClass.Heading7;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].NumberFormat = "%8";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].NumberStyle = WdListNumberStyle.wdListNumberStyleNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].NumberPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].TabPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].ResetOnHigher = 7;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[8].LinkedStyle = GlobalEnumClass.Heading8;

                //_withList = appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9];
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].NumberFormat = "%9.";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].NumberStyle = WdListNumberStyle.wdListNumberStyleLowercaseRoman;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].NumberPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].TextPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].TabPosition = appWord.InchesToPoints(0);
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].ResetOnHigher = 8;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].StartAt = 1;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Bold = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Italic = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.AllCaps = (int)GlobalEnumClass.AllEnum.False;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Underline = (WdUnderline)WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Size = (float)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Animation = (WdAnimation)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.DoubleStrikeThrough = (int)WdConstants.wdUndefined;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Name = "";
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].Font.Underline = WdUnderline.wdUnderlineNone;
                appWord.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[6].ListLevels[9].LinkedStyle = GlobalEnumClass.Heading9;
            }
            catch (Exception ex)
            {
                sbTrace.Clear();
                sbTrace.AppendLine("Exception" + ex);
                Logger.SaveLoggerTrace(sbTrace);
                Logger.LogWriter(ex.StackTrace);
            }
            finally
            {
                sbTrace.Clear();
                sbTrace.AppendLine("End");
                Logger.SaveLoggerTrace(sbTrace);
            }
        }
    }
}
