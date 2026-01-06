using Markdig;
using Markdig.Syntax;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Markdig.Syntax.Inlines;
namespace Function
{
	public static class MarkdownToRtfConverter
	{
		static Regex Color = new Regex("#[0-9A-Fa-f]{6}\\s");
		static Regex Table = new Regex("^(\\|.*\\|)$");
		static Dictionary<string, int> ColorDictionary;//其实可能有多个string指向同一个索引。转换成Color会多一步，不如直接拿string映射。因为多对一，所以不得不用Dictionary
		/// <summary>
		/// 将Color转换为Hex颜色代码，这样就可以用于显示颜色了。
		/// </summary>
		/// <param name="c"></param>
		/// <returns></returns>
		public static string ColorToHex(Color c)
		{
			return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
		}
		/// <summary>
		/// Main entry point: converts Markdown text to RTF.
		/// </summary>
		/// <param name="markdown">A string containing Markdown markup.</param>
		/// <returns>An RTF-formatted string.</returns>
		public static string Convert(string markdown)
		{
			// 1) Parse Markdown into a syntax tree
			var pipeline = new MarkdownPipelineBuilder().Build();
			var document = Markdig.Markdown.Parse(markdown, pipeline);

			// 2) Start building the RTF string
			var rtfBuilder = new StringBuilder();

			// Basic RTF header
			rtfBuilder.AppendLine(@"{\rtf1\ansi\deff0");
			//rtfBuilder.AppendLine(@"{\fonttbl{\f0 Arial;}}");

			// You can add more RTF preamble here (font tables, color tables, etc.)

			//制作颜色表
			ColorDictionary = new Dictionary<string, int>();
			List<Color> ColorList = new List<Color>();
			ColorList.Add(System.Drawing.Color.Black);
			ColorDictionary.Add("#000000 ", 0);
			var ms = Color.Matches(markdown);
			foreach (Match m in ms)
			{
				var C = ColorTranslator.FromHtml(m.Value.Substring(0, 7));
				int i = ColorList.IndexOf(C);
				if (i < 0)
				{
					i = ColorList.Count;
					ColorList.Add(C);
				}
				ColorDictionary.Add(m.Value, i);
			}
			if (ColorDictionary.Count > 0)
			{
				rtfBuilder.AppendLine(@"{\colortbl");

				foreach (var a in ColorDictionary)
				{
					var c = ColorList[a.Value];
					rtfBuilder.AppendLine($@"\t\red{c.R}\green{c.G}\blue{c.B};");
				}
				/*
				{\colortbl
					\red0\green0\blue0;       % 黑色
					\red255\green0\blue0;     % 红色
					\red0\green255\blue0;     % 绿色
					\red0\green0\blue255;     % 蓝色
					\red255\green255\blue0;   % 黄色
				}
				*/
				rtfBuilder.AppendLine(@"}");
			}
			rtfBuilder.Append("\\cf0");
			// 3) Walk the document blocks
			foreach (var block in document)
			{
				ConvertBlock(rtfBuilder, block);
			}

			// Close the RTF document
			rtfBuilder.AppendLine("}");

			return rtfBuilder.ToString();
		}

		/// <summary>
		/// Handles Markdown heading blocks (e.g., # Heading1, ## Heading2, etc.).
		/// </summary>
		private static void ConvertHeadingBlock(StringBuilder rtf, HeadingBlock headingBlock)
		{
			// Map heading level to font size (completely arbitrary example)
			// RTF uses \fsN in half-points (e.g., \fs32 => 16pt)
			int[] headingSizes = { 30, 28, 26, 24, 22, 20 };
			int headingLevel = headingBlock.Level; // 1-based

			// Get a font size for the heading level, clamp if needed
			int fontSize = headingSizes[Math.Min(headingLevel, headingSizes.Length) - 1];

			rtf.Append($@"\pard\sa180\fs{fontSize} \b ");
			// Heading text:
			ConvertInline(rtf, headingBlock.Inline);
			// End bold, new line
			rtf.AppendLine(@"\b0\par");
		}

		/// <summary>
		/// Handles normal paragraphs.
		/// </summary>
		private static void ConvertParagraphBlock(StringBuilder rtf, ParagraphBlock paragraphBlock)
		{
			rtf.Append(@"\pard\sa180\fs20 ");// \sa指段落间距，单位都是半磅，1/20的磅。 \fs指字体大小，不过转换成pt需要/2。
											 // Convert inlines inside this paragraph
			ConvertInline(rtf, paragraphBlock.Inline);
			// End paragraph
			rtf.AppendLine(@"\par");
		}

		/// <summary>
		/// Handles Markdown lists (ordered or unordered).
		/// </summary>
		private static void ConvertListBlock(StringBuilder rtf, ListBlock listBlock)
		{
			bool isOrdered = listBlock.IsOrdered;

			foreach (var item in listBlock)
			{
				// Each list item is itself a ListItemBlock containing sub-blocks.
				if (item is ListItemBlock listItemBlock)
				{
					// Start the bullet or number
					string prefix = isOrdered
						? $"{listItemBlock.Order}. "   // e.g., "1. ", "2. ", etc.
						: @"\bullet ";                // or just a bullet symbol, e.g. \bullet

					rtf.Append(@"\pard\sa100\fs20 ");
					rtf.Append(prefix);
					//rtf.Append(" ");

					// Convert each sub-block inside this list item
					foreach (var subBlock in listItemBlock)
					{
						switch (subBlock)
						{
							case ParagraphBlock subParagraph:
								ConvertInline(rtf, subParagraph.Inline);
								break;

							// Extend if you have nested lists, code, etc.
							default:
								break;
						}
					}

					// End list item
					rtf.AppendLine(@"\par");
				}
			}
		}

		/// <summary>
		/// Handles thematic breaks (horizontal rules) like '---'.
		/// </summary>
		private static void ConvertThematicBreakBlock(StringBuilder rtf, ThematicBreakBlock hrBlock)
		{
			rtf.AppendLine(@"\pard\brdrb\brdrs\brdrw10\brsp20\par");
		}

		/// <summary>
		/// 处理代码段
		/// </summary>
		private static void ConvertFencedCodeBlock(StringBuilder rtf, FencedCodeBlock block)
		{
			rtf.Append(@"\pard\sa0\fs18 ");// \sa指段落间距，单位都是半磅，1/20的磅。 \fs指字体大小，不过转换成pt需要/2。
			foreach (var a in block.Lines)
			{
				rtf.AppendLine(EscapeRtf(a.ToString()) + "\\line");//这个EscapeRtf真好啊，帮大忙了
			}
			rtf.AppendLine(@"\par");
		}

		/// <summary>
		/// Dispatches block conversion so it can be called recursively.
		/// </summary>
		private static void ConvertBlock(StringBuilder rtf, Block block)
		{
			switch (block)
			{
				case HeadingBlock headingBlock:
					ConvertHeadingBlock(rtf, headingBlock);
					break;

				case ParagraphBlock paragraphBlock:
					ConvertParagraphBlock(rtf, paragraphBlock);
					break;

				case ListBlock listBlock:
					ConvertListBlock(rtf, listBlock);
					break;

				case ThematicBreakBlock thematicBreakBlock:
					ConvertThematicBreakBlock(rtf, thematicBreakBlock);
					break;

				case QuoteBlock quoteBlock:
					ConvertQuoteBlock(rtf, quoteBlock);
					break;

				case FencedCodeBlock fencedCodeBlock:
					{
						ConvertFencedCodeBlock(rtf, fencedCodeBlock);
						break;
					}
				default:
					// Unhandled block type; extend as needed
					break;
			}
		}

		/// <summary>
		/// Handles block quotes.
		/// </summary>
		private static void ConvertQuoteBlock(StringBuilder rtf, QuoteBlock quoteBlock)
		{
			foreach (var subBlock in quoteBlock)
			{
				if (subBlock is ParagraphBlock paragraph)
				{
					rtf.Append(@"\pard\li300\sa180\fs20 ");
					ConvertInline(rtf, paragraph.Inline);
					rtf.AppendLine(@"\par");
				}
				else
				{
					ConvertBlock(rtf, subBlock);
				}
			}
		}
		static void DrawTable(StringBuilder rtf, List<string> ls)
		{
			if (ls.Count >= 3)
			{
				if (Regex.IsMatch(ls[1], "^((\\|-+)+\\|)$")) //第二行是分割线
				{
					ls.RemoveAt(1);
					//开始绘制
					List<int> Max = new List<int>();
					foreach (var s in ls)
					{
						var ss = s.Substring(1, s.Length - 2).Split('|');
						for (int i = 0; i < ss.Length; i++) ss[i] = ss[i].Trim();
						if (Max.Count == 0) for (int i = 0; i < ss.Length; i++) Max.Add(MeasureConsoleStringWidth(ss[i]));
						else for (int i = 0; i < ss.Length; i++) if (Max[i] < MeasureConsoleStringWidth(ss[i])) Max[i] = MeasureConsoleStringWidth(ss[i]);
						//需要字符串测量函数
					}
					//得到了每列的最大宽度
					bool First = true;
					foreach (var s in ls)
					{
						rtf.Append("\\pard");
						if (First)
						{
							rtf.Append("\\b ");//加粗
						}
						var ss = s.Substring(1, s.Length - 2).Split('|');
						for (int i = 0; i < ss.Length; i++) ss[i] = ss[i].Trim();
						for (int i = 0; i < ss.Length; i++)
						{
							rtf.Append(ss[i]);
							int l = MeasureConsoleStringWidth(ss[i]);
							rtf.Append(new string(' ', Max[i] - l + 1));//多留一格
						}
						if (First)
						{
							First = false;
							rtf.Append("\\b0 ");
						}
						rtf.Append("\\par");
					}
				}
				//也许是分成两个表格了？
				ls.RemoveAt(1);//去掉分割线那一行

				//else if (TableMode.Count == 1)
				//{
				//	if (Regex.IsMatch(s, "^((\\|-+)+\\|)$"))
				//	{
				//		rtf.Append("");
				//	}
				//}
			}
			else
			{
				foreach (var a in ls)
				{
					rtf.Append(EscapeRtf(a));
					rtf.Append(@"\par \pard\sa180\fs20 ");//与调用这里的一致。
				}
			}
			ls.Clear();
		}
		/// <summary>
		/// Recursively handles inlines (bold, italic, underline, etc.) in a Markdig Inline container.
		/// 爆改了，因为要处理表格，sorry, bro
		/// 因为改成迭代器了，如果不遍历元素就会不执行。所以调用的地方不出意外要用xxx.Any()
		/// </summary>
		private static void ConvertInline(StringBuilder rtf, ContainerInline containerInline, string prefix = "")
		{
			List<string> ls = new List<string>();
			bool IgnoreNextLine = false;
			foreach (var inline in containerInline)
			{
				switch (inline)
				{
					case EmphasisInline emphasisInline:
						HandleEmphasis(rtf, emphasisInline);
						break;

					case LineBreakInline lineBreakInline:
						// Soft line break or hard line break?
						// For simplicity, just do a line break.
						if (!IgnoreNextLine)
						{
							rtf.Append(@"\line ");
							IgnoreNextLine = false;
						}
						break;

					case CodeInline codeInline:
						// For code inline, you might do a monospace font or something else
						rtf.Append(@"\f1 "); // e.g., a monospace font
						rtf.Append(EscapeRtf(codeInline.Content));
						rtf.Append(@"\f0 ");
						break;

					case HtmlInline htmlInline:
						// Could try to interpret inline HTML, or just skip/escape
						rtf.Append(EscapeRtf(htmlInline.Tag));
						break;

					case LinkInline linkInline:
						// A link might show as underlined text + possibly a hidden URL
						// This is just a simplistic representation
						rtf.Append(@"\ul ");
						rtf.Append(EscapeRtf(prefix));
						rtf.Append(EscapeRtf(linkInline.Title ?? linkInline.Url));
						rtf.Append(@"\ulnone ");
						break;

					case LiteralInline literalInline://表格也要被归到这里了……硬写吧
						var s = literalInline.Content.ToString();
						bool HasColor = Color.IsMatch(s);
						if (HasColor)
						{
							var match = Color.Match(s);
							s = s.Remove(0, match.Value.Length);//#123456
							rtf.Append($"\\cf{ColorDictionary[match.Value]} ");//This is \cf2 red \cf0 text.  
						}
						if (Table.IsMatch(s))
						{
							ls.Add(s);
							IgnoreNextLine = true;
						}
						else
						{
							if (ls.Count >= 3) DrawTable(rtf, ls);
							else
							{
								DrawTable(rtf, ls);
								//foreach (var a in ls)
								//{
								//	rtf.Append(EscapeRtf(a));
								//	rtf.Append(@"\par \pard\sa180\fs20 ");//与调用这里的一致。
								//}
								//ls.Clear();
								IgnoreNextLine = false;
							}
							rtf.Append(EscapeRtf(s));//把每一行显示出来
						}
						if (HasColor) rtf.Append($" \\cf0 ");//恢复原始颜色
						break;
					default:
						// Not handled; no-op
						break;
				}
			}
			if (ls.Count > 0)
				DrawTable(rtf, ls);
		}

		/// <summary>
		/// Handles emphasis inlines (e.g., *italic*, **bold**, ***bold+italic***, etc.).
		/// We also interpret underscores as underline in this example.
		/// </summary>
		private static void HandleEmphasis(StringBuilder rtf, EmphasisInline emphasisInline)
		{
			// Markdig uses DelimiterChar = '*' for bold/italic, '_' is also possible
			bool isItalic = (emphasisInline.DelimiterChar == '*' && emphasisInline.DelimiterCount == 1)
							|| (emphasisInline.DelimiterChar == '_' && emphasisInline.DelimiterCount == 1);

			bool isBold = (emphasisInline.DelimiterChar == '*' && emphasisInline.DelimiterCount == 2);
			bool isUnderline = (emphasisInline.DelimiterChar == '_' && emphasisInline.DelimiterCount == 2);

			// For triple *** or ___, Markdig generally splits it into nested emphasis inlines
			// but you could handle combined styles if desired.

			// Start tags
			if (isBold) rtf.Append(@"\b ");
			if (isItalic) rtf.Append(@"\i ");
			if (isUnderline) rtf.Append(@"\ul ");

			// Recursively process the content inside the emphasis
			ConvertInline(rtf, emphasisInline);

			// End tags (reverse order)
			if (isUnderline) rtf.Append(@" \ulnone ");
			if (isItalic) rtf.Append(@" \i0 ");
			if (isBold) rtf.Append(@" \b0 ");
		}

		/// <summary>
		/// RTF is sensitive to certain special characters. Escape them here.
		/// </summary>
		private static string EscapeRtf(string text)
		{
			if (string.IsNullOrEmpty(text))
				return string.Empty;

			// Replace backslash, curly braces
			var sb = new StringBuilder();
			foreach (char c in text)
			{
				switch (c)
				{
					case '\\':
						sb.Append("\\\\");
						break;
					case '{':
						sb.Append("\\{");
						break;
					case '}':
						sb.Append("\\}");
						break;
					// Convert newline to \line or \par if desired, 
					// but let's do it in the inline logic instead
					default:
						sb.Append(c);
						break;
				}
			}//感觉能优化啊。

			return sb.ToString();
		}

		/// <summary>
		/// 测量一段的宽度，注意需要用等宽字体。
		/// </summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public static int MeasureConsoleStringWidth(string text)
		{
			int width = 0;
			int i = 0;
			//if (char.IsHighSurrogate(c) || char.IsLowSurrogate(c) || char.IsSurrogatePair(c))

			while (i < text.Length)
			{
				char c = text[i];
				if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
				{
					// 代理对（如某些 emoji）
					width += 2; // 假设代理对占用 2 个字符宽度
					i += 2; // 跳过代理对
				}
				else if (char.IsSurrogate(c))
				{
					// 单独的代理字符（理论上不应出现，但处理异常情况）
					width += 1; // 假设占用 1 个字符宽度
					i++;
				}
				else if (char.GetUnicodeCategory(c).ToString().StartsWith("Other"))
				{
					// 汉字等宽字符
					width += 2;
					i++;
				}
				else
				{
					// 英文、数字、符号等窄字符
					width += 1;
					i++;
				}
			}
			return width;
		}
	}
}
