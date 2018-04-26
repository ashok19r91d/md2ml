using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;

namespace Md2Ml
{
	public class Md2MlEngine : IDisposable
	{
		WordprocessingDocument package;
		MainDocumentPart mainDocumentPart;
		Document document;
		Body body;

		public void CreateDocument(string templateName, string fileName)
		{
			if (File.Exists(fileName))
				File.Delete(fileName);
			File.Copy(templateName, fileName);

			package = WordprocessingDocument.Open(fileName, true);
			mainDocumentPart = package.MainDocumentPart;
			document = mainDocumentPart.Document;
			body = document.Body;
			body.RemoveAllChildren<Paragraph>();
		}

		public Paragraph CreateParagraph()
		{
			var Para = new Paragraph();
			body.Append(Para);
			return Para;
		}
		public Paragraph CreateParagraph(string paragraphStyleName)
		{
			if (!(paragraphStyleName == null || paragraphStyleName == ""))
			{
				ParagraphProperties paraProp = new ParagraphProperties();
				paraProp.Append(new ParagraphStyleId() { Val = paragraphStyleName });
				var para = new Paragraph();
				para.Append(paraProp);
				body.Append(para);
				return para;
			}
			return CreateParagraph();
		}
		public Paragraph CreateParagraph(JustificationValues alignment)
		{
			if (alignment != JustificationValues.Left)
			{
				ParagraphProperties paraProp = new ParagraphProperties();
				paraProp.Append(new Justification() { Val = alignment });
				var para = CreateParagraph();
				para.Append(paraProp);
				return para;
			}
			return CreateParagraph();
		}
		public Paragraph CreateParagraph(ParaProperties properties)
		{
			ParagraphProperties paraProp = new ParagraphProperties();
			if (properties.StyleName != null) paraProp.Append(new ParagraphStyleId() { Val = properties.StyleName });
			if (properties.Alignment != JustificationValues.Left) paraProp.Append(new Justification() { Val = properties.Alignment });
			Indentation ind = new Indentation();
			if (properties.FirstLineIndent == 0) ind.FirstLine = (properties.FirstLineIndent * 567).ToString("n0").Replace(",", "");
			if (properties.LeftIndent == 0) ind.FirstLine = (properties.LeftIndent * 567).ToString("n0").Replace(",", "");
			if (properties.RightIndent == 0) ind.FirstLine = (properties.RightIndent * 567).ToString("n0").Replace(",", "");
			var para = CreateParagraph(); para.Append(paraProp); return para;
		}

		public Paragraph CreateNonBodyParagraph() => new Paragraph();
		public Paragraph CreateNonBodyParagraph(string paragraphStyleName)
		{
			if (!(paragraphStyleName == null || paragraphStyleName == ""))
			{
				ParagraphProperties paraProp = new ParagraphProperties();
				paraProp.Append(new ParagraphStyleId() { Val = paragraphStyleName });
				var para = CreateNonBodyParagraph();
				para.Append(paraProp);
				return para;
			}
			return CreateNonBodyParagraph();
		}
		public Paragraph CreateNonBodyParagraph(JustificationValues alignment)
		{
			if (alignment != JustificationValues.Left)
			{
				ParagraphProperties paraProp = new ParagraphProperties();
				paraProp.Append(new Justification() { Val = alignment });
				var para = CreateNonBodyParagraph();
				para.Append(paraProp);
				return para;
			}
			return CreateNonBodyParagraph();
		}
		public Paragraph CreateNonBodyParagraph(ParaProperties properties)
		{
			ParagraphProperties paraProp = new ParagraphProperties();
			if (properties.StyleName != null) paraProp.Append(new ParagraphStyleId() { Val = properties.StyleName });
			if (properties.Alignment != JustificationValues.Left) paraProp.Append(new Justification() { Val = properties.Alignment });
			Indentation ind = new Indentation();
			if (properties.FirstLineIndent == 0) ind.FirstLine = (properties.FirstLineIndent * 567).ToString("n0").Replace(",", "");
			if (properties.LeftIndent == 0) ind.FirstLine = (properties.LeftIndent * 567).ToString("n0").Replace(",", "");
			if (properties.RightIndent == 0) ind.FirstLine = (properties.RightIndent * 567).ToString("n0").Replace(",", "");
			var para = CreateNonBodyParagraph(); para.Append(paraProp); return para;
		}

		public void WriteText(string text, FontProperties fontProperties) => WriteText(CreateParagraph(), text, fontProperties);
		public void WriteText(Paragraph paragraph, string text, FontProperties fontProperties)
		{
			Run run = new Run();
			RunProperties rp = new RunProperties();
			if (fontProperties.StyleName != null)
				rp.Append(new RunStyle() { Val = fontProperties.StyleName });
			if (fontProperties.FontName != null)
				rp.Append(new RunFonts() { ComplexScript = fontProperties.FontName, Ascii = fontProperties.FontName, HighAnsi = fontProperties.FontName });
			else if (fontProperties.UseTemplateHeadingFont)
				rp.Append(new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, ComplexScriptTheme = ThemeFontValues.MajorHighAnsi });
			if (fontProperties.FontSize != null) { rp.Append(new FontSize() { Val = fontProperties.FontSize }); }
			if (fontProperties.Bold) rp.Append(new Bold());
			if (fontProperties.Italic) rp.Append(new Italic());
			if (fontProperties.Underline != UnderlineValues.None) rp.Append(new Underline() { Val = fontProperties.Underline });
			if (fontProperties.Strikeout) rp.Append(new Strike());
			if (fontProperties.WriteAs != VerticalPositionValues.Baseline) rp.Append(new VerticalTextAlignment() { Val = fontProperties.WriteAs });
			if (fontProperties.UseThemeColor) rp.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { ThemeColor = fontProperties.ThemeColor });
			else if (fontProperties.Color != null) rp.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = string.Format("#{0:X2}{1:X2}{2:X2}", fontProperties.Color.Value.R, fontProperties.Color.Value.G, fontProperties.Color.Value.B) });
			run.Append(rp);
			run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
			paragraph.Append(run);
		}
		public void WriteText(string text) => WriteText(CreateParagraph(), text);
		public void WriteText(Paragraph paragraph, string text) => WriteText(paragraph, text, new FontProperties());
		public void WriteMdText(string text) => MarkdownStringParser.Parse(this, text);
		public void WriteMdText(Paragraph para, string text) => MarkdownStringParser.Parse(this, text);

		public void WriteWebsiteHyperlink(string text, Uri uri, string relationshipId) => WriteWebsiteHyperlink(CreateParagraph("Hyperlink"), text, uri, relationshipId);
		public void WriteWebsiteHyperlink(Paragraph paragraph, string text, Uri uri, string relationshipId)
		{
			Hyperlink link = new Hyperlink() { Id = relationshipId };
			Run run = new Run();
			RunProperties runProperties = new RunProperties();
			RunStyle runStyle = new RunStyle() { Val = "Hyperlink" };
			runProperties.Append(runStyle);
			Text text1 = new Text() { Text = text };
			run.Append(runProperties);
			run.Append(text1);
			link.Append(run);
			paragraph.Append(link);
			mainDocumentPart.AddHyperlinkRelationship(uri, true, relationshipId);
		}

		public void MarkdownNumberedList(Md2MlEngine core, List<string> bulletedItems) => MarkdownNumberedList(core, bulletedItems, "ListParagraph");
		public void MarkdownNumberedList(Md2MlEngine core, List<string> bulletedItems, string paragraphStyle)
		{
			foreach (var item in bulletedItems)
			{
				Paragraph paragraph1 = CreateParagraph(paragraphStyle);
				NumberingProperties numberingProperties1 = new NumberingProperties();
				NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
				NumberingId numberingId1 = new NumberingId() { Val = 2 };
				numberingProperties1.Append(numberingLevelReference1);
				numberingProperties1.Append(numberingId1);
				paragraph1.ParagraphProperties.Append(numberingProperties1);
				MarkdownStringParser.FormatText(core, paragraph1, item, new FontProperties());
			}
		}
		public void NumberedList(List<string> bulletedItems) => NumberedList(bulletedItems, "ListParagraph");
		public void NumberedList(List<string> bulletedItems, string paragraphStyle)
		{
			foreach (var item in bulletedItems)
			{
				Paragraph paragraph1 = CreateParagraph(paragraphStyle);
				NumberingProperties numberingProperties1 = new NumberingProperties();
				NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
				NumberingId numberingId1 = new NumberingId() { Val = 2 };
				numberingProperties1.Append(numberingLevelReference1);
				numberingProperties1.Append(numberingId1);
				paragraph1.ParagraphProperties.Append(numberingProperties1);
				Run run1 = new Run();
				Text text1 = new Text();
				text1.Text = item;
				run1.Append(text1);
				paragraph1.Append(run1);
			}
		}

		public void BulletedList(List<string> bulletedItems) => BulletedList(bulletedItems, "ListParagraph");
		public void BulletedList(List<string> bulletedItems, string paragraphStyle)
		{
			foreach (var item in bulletedItems)
			{
				Paragraph paragraph1 = CreateParagraph(paragraphStyle);
				NumberingProperties numberingProperties1 = new NumberingProperties();
				NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
				NumberingId numberingId1 = new NumberingId() { Val = 1 };
				numberingProperties1.Append(numberingLevelReference1);
				numberingProperties1.Append(numberingId1);
				paragraph1.ParagraphProperties.Append(numberingProperties1);
				Run run1 = new Run();
				Text text1 = new Text();
				text1.Text = item;
				run1.Append(text1);
				paragraph1.Append(run1);
			}
		}
		public void MarkdownBulletedList(Md2MlEngine core, List<string> bulletedItems) => MarkdownBulletedList(core, bulletedItems, "ListParagraph");
		public void MarkdownBulletedList(Md2MlEngine core, List<string> bulletedItems, string paragraphStyle)
		{
			foreach (var item in bulletedItems)
			{
				Paragraph paragraph1 = CreateParagraph(paragraphStyle);
				NumberingProperties numberingProperties1 = new NumberingProperties();
				NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
				NumberingId numberingId1 = new NumberingId() { Val = 1 };
				numberingProperties1.Append(numberingLevelReference1);
				numberingProperties1.Append(numberingId1);
				paragraph1.ParagraphProperties.Append(numberingProperties1);
				MarkdownStringParser.FormatText(core, paragraph1, item, new FontProperties());
			}
		}

		public void AddPageBreak() => AddPageBreak(CreateParagraph());
		public void AddPageBreak(Paragraph paragraph)
		{
			var Run = new Run();
			Break Break = new Break() { Type = BreakValues.Page };
			Run.Append(Break);
			paragraph.Append(Run);
		}

		public void StartBookmark(Paragraph paragraph, string bookmarkName, string bookmarkId) => paragraph.Append(new BookmarkStart() { Name = bookmarkName, Id = bookmarkId });
		public void EndBookmark(Paragraph paragraph, string bookmarkId) => paragraph.Append(new BookmarkEnd() { Id = bookmarkId });

		public Table CreateTable(int cols)
		{
			Table table = new Table();
			TableProperties tableProperties = new TableProperties();
			TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
			TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
			TableLook tableLook = new TableLook() { Val = "04A0" };
			tableProperties.Append(tableStyle);
			tableProperties.Append(tableWidth);
			tableProperties.Append(tableLook);
			TableGrid tableGrid = new TableGrid();
			for (int i = 0; i < cols; i++)
			{
				GridColumn gridColumn = new GridColumn() { Width = (10216 / cols).ToString("n0").Replace(",", "") };
				tableGrid.Append(gridColumn);
			}
			table.Append(tableProperties);
			table.Append(tableGrid);
			body.Append(table);
			return table;
		}
		public Table CreateTable(int cols, int[] widths)
		{
			Table table = new Table();
			TableProperties tableProperties = new TableProperties();
			TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
			TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
			TableLook tableLook = new TableLook() { Val = "04A0" };
			tableProperties.Append(tableStyle);
			tableProperties.Append(tableWidth);
			tableProperties.Append(tableLook);
			TableGrid tableGrid = new TableGrid();
			for (int i = 0; i < cols; i++)
			{
				GridColumn gridColumn = new GridColumn() { Width = (widths[i] * 102.16).ToString("n0").Replace(",", "") };
				tableGrid.Append(gridColumn);
			}
			table.Append(tableProperties);
			table.Append(tableGrid);
			body.Append(table);
			return table;
		}
		public void AddTableRow(Table table, List<string> values, int[] widths)
		{
			TableRow tableRow = new TableRow();
			int i = 0;
			foreach (var value in values)
			{
				TableCell tableCell = new TableCell();
				TableCellProperties tableCellProperties = new TableCellProperties();
				TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = widths[i++].ToString() };
				tableCellProperties.Append(tableCellWidth);
				var para = CreateNonBodyParagraph();
				WriteText(para, value);
				tableCell.Append(tableCellProperties);
				tableCell.Append(para);
				tableRow.Append(tableCell);
			}
			table.Append(tableRow);
		}
		public void AddTableRow(Table table, List<Paragraph> paras, int[] widths)
		{
			TableRow tableRow = new TableRow();
			int i = 0;
			foreach (var para in paras)
			{
				TableCell tableCell = new TableCell();
				TableCellProperties tableCellProperties = new TableCellProperties();
				TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = widths[i++].ToString() };
				tableCellProperties.Append(tableCellWidth);
				tableCell.Append(tableCellProperties);
				tableCell.Append(para);
				tableRow.Append(tableCell);
			}
			table.Append(tableRow);
		}
		public void AddTableRow(Table table, List<string> values)
		{
			List<int> width = new List<int>();
			for (var i = 0; i < values.Count; i++)
				width.Add(int.Parse((10216 / values.Count).ToString("n0").Replace(",", "")));

			AddTableRow(table, values, width.ToArray());
		}
		public void AddTableRow(Table table, List<Paragraph> paras)
		{
			List<int> width = new List<int>();
			width.Add(int.Parse((10216 / paras.Count).ToString("n0")));
			AddTableRow(table, paras, width.ToArray());
		}

		public void PageSetup(PageSize pageSize) => throw new NotImplementedException();

		public void AddImage(Stream image)
		{
			ImagePart imagePart = mainDocumentPart.AddImagePart("image/png");
			imagePart.FeedData(image);
			AddImageToBody(package, mainDocumentPart.GetIdOfPart(imagePart));
		}
		private static void AddImageToBody(WordprocessingDocument document, string relationshipId)
		{
			var element =
				 new Drawing(
					 new Wp.Inline(
						 new Wp.Extent() { Cx = 990000L, Cy = 792000L },
						 new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
						 new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
						 new Wp.NonVisualGraphicFrameDrawingProperties( new A.GraphicFrameLocks() { NoChangeAspect = true }),
						 new A.Graphic(
							 new A.GraphicData(
								 new Pic.Picture(
									 new Pic.NonVisualPictureProperties(
										 new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.Png" },
										 new Pic.NonVisualPictureDrawingProperties()),
									 new Pic.BlipFill(
										 new A.Blip(
											 new A.BlipExtensionList(
												 new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })
										 )
										 {
											 Embed = relationshipId,
											 CompressionState = A.BlipCompressionValues.Print
										 },
										 new A.Stretch(
											 new A.FillRectangle())),
									 new Pic.ShapeProperties(
										 new A.Transform2D(
											 new A.Offset() { X = 0L, Y = 0L },
											 new A.Extents() { Cx = 990000L, Cy = 792000L }),
										 new A.PresetGeometry( new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
							 )
							 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
					 )
					 {
						 DistanceFromTop = (UInt32Value)0U,
						 DistanceFromBottom = (UInt32Value)0U,
						 DistanceFromLeft = (UInt32Value)0U,
						 DistanceFromRight = (UInt32Value)0U,
						 EditId = "50D07946"
					 });

			// Append the reference to body, the element should be in a Run.
			document.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

		}

		public void Cleanup(OpenXmlElement element) => element.RemoveAllChildren();

		public void SaveDocument() => package.Save();
		public void SaveDocument(string fileName) => package.SaveAs(fileName);

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			SaveDocument();
			if (!disposedValue)
			{
				if (disposing)
				{
					package.Dispose();
				}
				package = null;
				disposedValue = true;
			}
		}
		public void Dispose() => Dispose(true);
		#endregion
	}
}
