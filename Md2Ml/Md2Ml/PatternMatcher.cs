using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Md2Ml
{
	class PatternMatcher
	{
		private static Dictionary<RunPattern, Regex> RunPatters = new Dictionary<RunPattern, Regex>()
		{
			{ RunPattern.Bold, new Regex(@"(?<!\*)\*\*([^\*].+?)\*\*") },
			{ RunPattern.Italic, new Regex(@"(?<!\*)\*([^\*].+?)\*") },
			{ RunPattern.BoldAndItalic, new Regex(@"(?<!\*)\*\*\*([^\*].+?)\*\*\*") },
			{ RunPattern.Link, new Regex(@"\[(.+?)\]\((.+)\)") },
			{ RunPattern.MonospaceOrCode, new Regex(@"`{1}([^`]+)`{1}") },
			{ RunPattern.Strikethrough, new Regex(@"~{2}(.*)~{2,}") },
			{ RunPattern.Tab, new Regex(@"\t(.*)") },
			{ RunPattern.Underline, new Regex(@"_{2}(.*)_{2,}") }
		};
		private static Dictionary<ParaPattern, Regex> ParagraphPatters = new Dictionary<ParaPattern, Regex>()
		{
			{ ParaPattern.CodeBlock, new Regex(@"[ ]{4}(.*)") },
			{ ParaPattern.Heading1, new Regex(@"^# (.*)") },
			{ ParaPattern.Heading2, new Regex(@"^## (.*)") },
			{ ParaPattern.Heading3, new Regex(@"^### (.*)") },
			{ ParaPattern.Image, new Regex(@"\!\[(.+?)\]\((.+)\)") },
			{ ParaPattern.OrderedList, new Regex(@"^[\d]\. (.*)") },
			{ ParaPattern.Quote, new Regex(@">{1} (.*)") },
			{ ParaPattern.Table, new Regex(@"\|(.*)\|") },
			{ ParaPattern.UnorderedList, new Regex(@"^[*+-] (.*)") },
			{ ParaPattern.CommanBlock, new Regex("(.*)") }
		};
		public static KeyValuePair<ParaPattern, string[]> GetParagraphType(string markdown)
		{
			foreach (var pattern in ParagraphPatters)
			{
				var regex = pattern.Value;
				if (!regex.IsMatch(markdown)) continue;
				return new KeyValuePair<ParaPattern, string[]>(pattern.Key, regex.Split(markdown));
			}
			throw new NotSupportedException();
		}
		public static bool HasPatterns(string markdown)
		{
			foreach (var pattern in RunPatters)
			{
				var regex = pattern.Value;
				if (!regex.IsMatch(markdown)) continue;
				return true;
			}
			return false;
		}
		public static KeyValuePair<RunPattern, string[]> GetPatternsAndNonPatternText(string markdown)
		{
			foreach (var pattern in RunPatters)
			{
				var regex = pattern.Value;
				if (!regex.IsMatch(markdown)) continue;
				return new KeyValuePair<RunPattern, string[]>(pattern.Key, regex.Split(markdown));
			}
			return new KeyValuePair<RunPattern, string[]>(RunPattern.PlainText, new string[] { markdown });
		}
	}
}