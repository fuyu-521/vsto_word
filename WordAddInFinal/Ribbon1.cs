using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Word;

namespace WordAddInFinal
{
    public partial class Ribbon1
    {




        private void ExportDocument(object sender, RibbonControlEventArgs e)
        {

            switch (e.Control.Id)
            {
                // 判断点击的按钮ID
                case "button1":
                    // 打开保存文件窗口
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        // 设置保存文件窗口的相关属性
                        saveFileDialog.Filter = "All File(*.*)|*.*";
                        saveFileDialog.DefaultExt = ".pdf";
                        saveFileDialog.RestoreDirectory = true;
                        // 在保存文件窗口中点击保存按钮
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            // 导出为PDF格式
                            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                                saveFileDialog.FileName,
                                Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        }
                    }
                    break;
                case "button2":
                    using (SaveFileDialog saveFileDialog2 = new SaveFileDialog())
                    {
                        saveFileDialog2.Filter = "All File(*.*)|*.*";
                        saveFileDialog2.DefaultExt = ".xps";
                        saveFileDialog2.RestoreDirectory = true;
                        if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                        {
                            // 导出为XPS格式
                            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                                saveFileDialog2.FileName,
                                Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatXPS);
                        }
                    }
                    break;
                default:
                    return;
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            button1.Click += new RibbonControlEventHandler(ExportDocument);
            button2.Click += new RibbonControlEventHandler(ExportDocument);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            app.Selection.WholeStory();



        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前打开的Word文档
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.PageSetup pageSetup = currentDocument.PageSetup;

            // 设置纸张大小为A4
            currentDocument.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;

            // 设置页边距
            pageSetup.TopMargin = currentDocument.Application.CentimetersToPoints(2.5f);
            pageSetup.BottomMargin = currentDocument.Application.CentimetersToPoints(2.0f);
            pageSetup.LeftMargin = currentDocument.Application.CentimetersToPoints(3.0f);
            pageSetup.RightMargin = currentDocument.Application.CentimetersToPoints(2.0f);
            
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "黑体";
            currentSelection.Font.Size = 16;
            currentSelection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "黑体";
            currentSelection.Font.Size = 14;
            currentSelection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            currentSelection.ParagraphFormat.CharacterUnitLeftIndent = 0;
            currentSelection.ParagraphFormat.CharacterUnitRightIndent = 0;
            currentSelection.ParagraphFormat.FirstLineIndent = 0;
            currentSelection.ParagraphFormat.LeftIndent = 0;
            currentSelection.ParagraphFormat.RightIndent = 0;
            currentSelection.ParagraphFormat.SpaceBefore = 12; // 0.5 行对应的磅数是 12 磅
            currentSelection.ParagraphFormat.SpaceAfter = 0;

        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "黑体";
            currentSelection.Font.Size = 12;
            currentSelection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            currentSelection.ParagraphFormat.SpaceBefore = 0; 
            currentSelection.ParagraphFormat.SpaceAfter = 0;

        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "黑体";
            currentSelection.Font.Size = 16;
            currentSelection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "宋体";
            currentSelection.Font.Size = 12;
            
            currentSelection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            currentSelection.ParagraphFormat.LineSpacing = currentDocument.Application.InchesToPoints(22.0f / 72.0f);
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "Times New Roman";
            currentSelection.Font.Size = 16;
            currentSelection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "黑体";
            currentSelection.Font.Size = 12;
            currentSelection.Font.Bold = 1;

        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "Times New Roman";
            currentSelection.Font.Size = 12;
            currentSelection.Font.Bold = 1;
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "Times New Roman";
            currentSelection.Font.Size = 12;

            currentSelection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            currentSelection.ParagraphFormat.LineSpacing = currentDocument.Application.InchesToPoints(22.0f / 72.0f);

        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 设置文本格式
            currentSelection.Font.Name = "宋体";
            currentSelection.Font.Name = "Times New Roman";
            currentSelection.Font.Size = 12;
            currentSelection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            currentSelection.ParagraphFormat.LineSpacing = 22f;
            currentSelection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            currentSelection.ParagraphFormat.SpaceBefore = 0;
            currentSelection.ParagraphFormat.SpaceAfter = 0;
        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;

            // 获取选定的文本对象
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            currentSelection.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档的选区
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 插入“目录”文本
            currentSelection.TypeText("目  录");

            // 获取刚插入的文本范围
            Word.Range range = currentSelection.Range;
            range.Start -= 4;
            range.End -= 0;

            // 设置文本属性
            range.Font.Name = "黑体";
            range.Font.Size = 16;
            
            range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

        }

        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档对象
            //Microsoft.Office.Interop.Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;


            // 插入目录
            //currentDocument.TablesOfContents.Add(currentDocument.Sections[1].Range, true, 1, 3);

            // 设置指定段落的大纲级别
            //currentDocument.Content.Paragraphs[2].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading1]);
            //currentDocument.Content.Paragraphs[4].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading2]);
            //currentDocument.Content.Paragraphs[7].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading2]);
            //currentDocument.Content.Paragraphs[10].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading2]);
            //currentDocument.Content.Paragraphs[12].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading1]);
            //currentDocument.Content.Paragraphs[14].Range.set_Style(currentDocument.Styles[WdBuiltinStyle.wdStyleHeading1]);

            // 更新目录
            //currentDocument.TablesOfContents[1].Update();
        }

        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档的选区
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;

            // 插入文本
            currentSelection.TypeText("参考文献");

            // 获取刚插入的文本范围
            Word.Range range = currentSelection.Range;
            range.Start -= 4;
            range.End -= 0;

            // 设置文本属性
            range.Font.Name = "黑体";
            range.Font.Size = 16;
            
            range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

        }

        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前文档的选区
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            // 设置文本格式
            currentSelection.Font.Name = "宋体";
            currentSelection.Font.Name = "Times New Roman";
            currentSelection.Font.Size = 12;
            currentSelection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            currentSelection.ParagraphFormat.LineSpacing = 22f;
            currentSelection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            currentSelection.ParagraphFormat.CharacterUnitLeftIndent = 0;
            currentSelection.ParagraphFormat.CharacterUnitRightIndent = 0;
            currentSelection.ParagraphFormat.FirstLineIndent = 0;
            currentSelection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void button18_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
