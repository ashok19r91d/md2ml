using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace Md2Ml
{
	public class FontProperties
	{
		public string StyleName = null;
		public string FontName = null;
		public string FontSize = null;
		public bool Bold = false;
		public bool Italic = false;
		public UnderlineValues Underline = UnderlineValues.None;
		public bool Strikeout = false;
		public VerticalPositionValues WriteAs = VerticalPositionValues.Baseline;
		public System.Drawing.Color? Color = null;
		public ThemeColorValues ThemeColor = ThemeColorValues.Text1;
		public bool UseThemeColor = false;
		public bool UseTemplateHeadingFont = false;
	}
}
