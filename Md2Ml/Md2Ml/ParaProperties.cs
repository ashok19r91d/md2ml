using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace Md2Ml
{
	public class ParaProperties
	{
		public decimal FirstLineIndent = 0;
		public decimal LeftIndent = 0;
		public decimal RightIndent = 0;
		public JustificationValues Alignment = JustificationValues.Both;
		public string StyleName = null;
	}
}
