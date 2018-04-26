using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Md2Ml
{
	class MarkdownStringParser
	{
		public static void Parse(Md2MlEngine engine, string mdText)
		{
			var lineAndPattern = new List<KeyValuePair<ParaPattern, string[]>>();
			var lines = mdText.Split('\n');
			foreach (var line in lines)
				lineAndPattern.Add(PatternMatcher.GetParagraphType(line));
			bool OrderedList = false;
			bool UnorderedList = false;
			bool Table = false;
			List<string> BulletItems = new List<string>();
			List<string> NumberItems = new List<string>();
			List<string> TableData = new List<string>();
			foreach (var line in lineAndPattern)
			{
				Paragraph para;
				switch (line.Key)
				{
					case ParaPattern.OrderedList:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (UnorderedList) { ProcessBullets(engine, BulletItems, false); UnorderedList = false; }
						OrderedList = true;
						NumberItems.Add(line.Value[1]);
						break;
					case ParaPattern.UnorderedList:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList) { ProcessBullets(engine, NumberItems, true); OrderedList = false; }
						UnorderedList = true;
						BulletItems.Add(line.Value[1]);
						break;
					case ParaPattern.Image:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph();
						if (line.Value[2].StartsWith("http://") || line.Value[2].StartsWith("https://"))
							engine.AddImage(new System.Net.WebClient().OpenRead(line.Value[2]));
						else
							engine.AddImage(System.IO.File.OpenRead(line.Value[2]));
						break;
					case ParaPattern.Table:
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						Table = true;
						TableData.Add(line.Value[1]);
						break;
					case ParaPattern.CodeBlock:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph(new ParaProperties() { StyleName = "CodeBlock" });
						engine.WriteText(para, line.Value[1]);
						break;
					case ParaPattern.Heading1:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph(new ParaProperties() { StyleName = "Heading1" });
						engine.WriteText(para, line.Value[1]);
						break;
					case ParaPattern.Heading2:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph(new ParaProperties() { StyleName = "Heading2" });
						engine.WriteText(para, line.Value[1]);
						break;
					case ParaPattern.Heading3:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph(new ParaProperties() { StyleName = "Heading3" });
						engine.WriteText(para, line.Value[1]);
						break;
					case ParaPattern.Quote:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph(new ParaProperties() { StyleName = "Quote" });
						engine.WriteText(para, line.Value[1]);
						break;
					case ParaPattern.CommanBlock:
					default:
						if (Table) { ProcessTable(engine, TableData); Table = false; }
						if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
						para = engine.CreateParagraph();
						FormatText(engine, para, line.Value[1], new FontProperties());
						//core.WriteText(para, line.Value[1]);
						break;
				}
			}
			if (Table) { ProcessTable(engine, TableData); Table = false; }
			if (OrderedList || UnorderedList) { ProcessBullets(engine, OrderedList ? NumberItems : BulletItems, OrderedList); OrderedList = false; UnorderedList = false; }
		}
		public static void ProcessBullets(Md2MlEngine core, List<string> bullets, bool ordered = false)
		{
			if (bullets.Count != 0)
			{
				if (ordered)
					core.MarkdownNumberedList(core, bullets);
				else
					core.MarkdownBulletedList(core, bullets);
				bullets.Clear();
			}
		}
		public static void FormatText(Md2MlEngine core, Paragraph paragraph, string markdown, FontProperties fontProperties)
		{
			var hasPattern = PatternMatcher.HasPatterns(markdown);
			while (hasPattern)
			{
				var s = PatternMatcher.GetPatternsAndNonPatternText(markdown);
				var count = s.Value.Count();
				var NewFontProperties = new FontProperties();
				switch (s.Key)
				{
					case RunPattern.BoldAndItalic:
						NewFontProperties.Bold = true;
						NewFontProperties.Italic = true;
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "***"), new FontProperties());
						break;
					case RunPattern.Bold:
						NewFontProperties.Bold = true;
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "**"), new FontProperties());
						break;
					case RunPattern.Italic:
						NewFontProperties.Italic = true;
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "*"), new FontProperties());
						break;
					case RunPattern.MonospaceOrCode:
						NewFontProperties.StyleName = "InlineCodeChar";
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "`"), new FontProperties());
						break;
					case RunPattern.Strikethrough:
						NewFontProperties.Strikeout = true;
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "~~"), new FontProperties());
						break;
					case RunPattern.Underline:
						NewFontProperties.Underline = UnderlineValues.Single;
						FormatText(core, paragraph, s.Value[0], new FontProperties());
						FormatText(core, paragraph, s.Value[1], NewFontProperties);
						FormatText(core, paragraph, FramePendingString(s.Value, "__"), new FontProperties());
						break;
				}
				return;
			}
			core.WriteText(paragraph, markdown, fontProperties);
		}
		private static string FramePendingString(string[] strs, string patten)
		{
			var str = "";
			for (int i = 2; i < strs.Count(); i++)
			{
				if (i % 2 == 0)
					str += strs[i];
				else
					str += patten + strs[i] + patten;
			}
			return str;
		}
		private static void ProcessTable(Md2MlEngine core, List<string> markdown)
		{
			var table = core.CreateTable(markdown.First().Trim('|').Split('|').Count());
			core.AddTableRow(table, markdown.First().Trim(new char[] { '|' }).Split('|').ToList());
			foreach (var data in markdown.Skip(2))
				core.AddTableRow(table, data.Trim(new char[] { '|' }).Split('|').ToList());
		}

	}
}
