using Markdig;
using Markdig.Syntax;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
namespace Function
{
	public static class MarkdownToRtfConverter
	{
		static Regex Color = new Regex("#[0-9A-Fa-f]{6}\\s");
		//static Regex Table = new Regex("^(\\|.*\\|)$");
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
		static int BaseFontSize = 20;
		/// <summary>
		/// Main entry point: converts Markdown text to RTF.
		/// </summary>
		/// <param name="markdown">A string containing Markdown markup.</param>
		/// <param name="baseFontSize">基础字体大小，为pt数的两倍。</param>
		/// <param name="FontColor">字体默认颜色，留空就是黑</param>
		/// <returns>An RTF-formatted string.</returns>
		public static string Convert(string markdown, int baseFontSize = 20, Color? FontColor = null)
		{
			BaseFontSize = baseFontSize;
			// 1) Parse Markdown into a syntax tree
			var pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();// 启用表格扩展

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
			if (FontColor == null)
			{
				ColorList.Add(System.Drawing.Color.Black);
				ColorDictionary.Add("#000000 ", 0);
			}
			else
			{
				ColorList.Add(FontColor.Value);
				ColorDictionary.Add(ColorToHex(FontColor.Value) + " ", 0);
			}
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
				if (!ColorDictionary.ContainsKey(m.Value)) ColorDictionary.Add(m.Value, i);
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
			Layer = -1;//用于处理嵌套的ListBlock
			NeedPar = false;
			foreach (var block in document)
			{
				ConvertBlock(rtfBuilder, block);
			}

			// Close the RTF document
			rtfBuilder.AppendLine("}");

			//还需要处理Emoji的代理对。
			/*
			高代理字符的编码值需要减去 0x10000，然后右移 10 位，加上 0xD800。
			低代理字符的编码值需要与 0x3FF 进行按位与操作，然后加上 0xDC00
			*/
			for (int i = rtfBuilder.Length - 1; i >= 0; i--)
			{
				if (char.IsSurrogate(rtfBuilder[i]))
				{
					if (i >= 1)
					{
						if (char.IsHighSurrogate(rtfBuilder[i - 1]) && char.IsLowSurrogate(rtfBuilder[i]))
						{
							// 计算 RTF 格式的 Unicode 编码
							int highSurrogateValue = rtfBuilder[i - 1];
							int lowSurrogateValue = rtfBuilder[i];
							rtfBuilder.Remove(i - 1, 2);
							rtfBuilder.Insert(i - 1, $"\\u{highSurrogateValue}?\\u{lowSurrogateValue}?");
							//这里其实取巧了。遇到的微笑【😊】在rtf中显示是\u-10179?\u-8700?，但还好RTF也认\u55357?\u56842?
							i--; // 跳过低代理字符
						}
					}
					else
					{
						throw new Exception("一位的代理？");
					}
				}
			}
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

			rtf.Append($@"\pard\sa180\fs{(int)(fontSize / 20.0 * BaseFontSize)} \b ");
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
			rtf.Append($@"\pard\sa180\fs{BaseFontSize} ");// \sa指段落间距，单位都是半磅，1/20的磅。 \fs指字体大小，不过转换成pt需要/2。
														  // Convert inlines inside this paragraph
			ConvertInline(rtf, paragraphBlock.Inline);
			// End paragraph
			rtf.AppendLine(@"\par");
		}
		static bool NeedPar = false;

		/// <summary>
		/// Handles Markdown lists (ordered or unordered).
		/// </summary>
		private static void ConvertListBlock(StringBuilder rtf, ListBlock listBlock)
		{
			bool isOrdered = listBlock.IsOrdered;
			if (Layer >= 0) rtf.AppendLine(@"\par");
			Layer++;
			foreach (var item in listBlock)
			{
				// Each list item is itself a ListItemBlock containing sub-blocks.
				if (item is ListItemBlock listItemBlock)
				{
					// Start the bullet or number
					if (NeedPar)
						rtf.AppendLine(@"\par");
					string prefix = isOrdered
						? $"{listItemBlock.Order}. "   // e.g., "1. ", "2. ", etc.
						: @"\bullet ";                // or just a bullet symbol, e.g. \bullet

					rtf.Append($@"\pard\sa100\li{360 * Layer}\fs{BaseFontSize} ");
					rtf.Append(prefix);
					//rtf.Append(" ");

					// Convert each sub-block inside this list item
					foreach (var subBlock in listItemBlock)
					{
						switch (subBlock)
						{
							case ParagraphBlock subParagraph:
								Console.WriteLine(new string('\t', Layer) + "ParagraphBlock");
								ConvertInline(rtf, subParagraph.Inline);//相邻两个这个中间需要换行……
								NeedPar = true;
								break;
							case ListBlock listBlock2:
								Console.WriteLine(new string('\t', Layer) + "ListBlock");
								NeedPar = false;
								//这里缺少一个分行
								//缺少缩进，也就是需要记录listBlock的嵌套深度。html代码本身没问题，需要看看/par加在什么地方。
								ConvertBlock(rtf, listBlock2);//疑似套娃
								break;
							// Extend if you have nested lists, code, etc.
							default:
								break;
						}
					}

					// End list item
					//if(Layer==0)
				}
			}
			//rtf.AppendLine(@"\par");

			Layer--;
			if (Layer < 0)
			{
				rtf.AppendLine(@"\par");
				NeedPar = false;
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
			rtf.Append($@"\pard\sa0\fs{BaseFontSize * 9 / 10} ");// \sa指段落间距，单位都是半磅，1/20的磅。 \fs指字体大小，不过转换成pt需要/2。
			foreach (var a in block.Lines)
			{
				rtf.AppendLine(EscapeRtf(a.ToString()) + "\\line");//这个EscapeRtf真好啊，帮大忙了
			}
			rtf.AppendLine(@"\par");
		}
		static int Layer = 0;
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
				case Table Rows://什么嘛，原来Markdig支持表格
					{
						List<int> Max = new List<int>();
						List<(string text, string rtf, int length)[]> SSS = new List<(string text, string rtf, int length)[]>();//记录每一格的文字长度、文字原始数据以及带格式数据（rtf）
						for (int j = 0; j < Rows.Count; j++)//总感觉循环条件怪怪的。
						{
							var Columns = Rows[j] as TableRow;
							if (Max.Count == 0) for (int i = 0; i < Columns.Count; i++) Max.Add(0);
							(string text, string rtf, int length)[] ss = new (string text, string rtf, int length)[Columns.Count];
							for (int k = 0; k < Columns.Count; k++)
							{
								var cell = (TableCell)Columns[k];
								StringBuilder Text = new StringBuilder();//存放纯文本，用于计算每列的最大宽度
								StringBuilder TempRtf = new StringBuilder();//存放带格式的文本
								foreach (LeafBlock data in cell)//cell里会有多个data吗？
								{
									var containerInline = data.Inline;
									var literalInline = containerInline.FirstChild as LiteralInline;
									if (literalInline.NextSibling != null)
									{
										//可能有格式符
										ConvertInline(TempRtf, containerInline, Text);
										//需要得到文字宽度，计算每列的最大列宽
									}
									//要考虑格式符，.NextSibling 
									else
									{
										Text.Append(literalInline.Content);
									}
								}
								var l = MeasureConsoleStringWidth(Text.ToString());
								ss[k] = (Text.ToString(), TempRtf.ToString(), l);
								if (Max[k] < l)
									Max[k] = l;
							}
							SSS.Add(ss);
						}
						//接下来是输出……
						foreach (var ss in SSS)
						{
							rtf.Append("\\pard ");
							for (int i = 0; i < ss.Length; i++)
							{
								string s;
								if (string.IsNullOrEmpty(ss[i].rtf)) rtf.Append(s = ss[i].text);
								else rtf.Append(s = ss[i].rtf);
								rtf.Append(new string(' ', Max[i] - ss[i].length));
							}
							rtf.Append("\\par");
						}
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
					rtf.Append($@"\pard\li300\sa180\fs{BaseFontSize} ");
					ConvertInline(rtf, paragraph.Inline);
					rtf.AppendLine(@"\par");
				}
				else
				{
					ConvertBlock(rtf, subBlock);
				}
			}
		}
		/// <summary>
		/// Recursively handles inlines (bold, italic, underline, etc.) in a Markdig Inline container.
		/// </summary>
		private static void ConvertInline(StringBuilder rtf, ContainerInline containerInline, StringBuilder Text = null, string prefix = "")
		{
			//List<string> ls = new List<string>();
			bool IgnoreNextLine = false;
			foreach (var inline in containerInline)
			{
				switch (inline)
				{
					case EmphasisInline emphasisInline:
						HandleEmphasis(rtf, emphasisInline, Text);
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
													 //rtf.Append(EscapeRtf(literalInline.Content.ToString()));//原始
						var s = literalInline.Content.ToString();
						if (Text != null) Text.Append(s);
						bool HasColor = Color.IsMatch(s);
						if (HasColor)
						{
							var match = Color.Match(s);
							s = s.Remove(0, match.Value.Length);//#123456
							rtf.Append($"\\cf{ColorDictionary[match.Value]} ");//This is \cf2 red \cf0 text.  
						}
						//当表格格式为| **abc** |时，s只是|。后续会进入emphasisInline，配置格式后再回到literalInline……这无疑是一个挑战。
						//两种思路
						//一种是提前把所有表格部分替换成特殊的标记，在程序里单独处理完毕后塞回来。
						//另一种是检测|，以|作为表格一行的起点和分割。适用面更广，也更麻烦……
						//然而MarkDig支持表格
						rtf.Append(EscapeRtf(s));//把每一行显示出来
						if (HasColor) rtf.Append($" \\cf0 ");//恢复原始颜色
						break;
					default:
						// Not handled; no-op
						break;
				}
			}
			//if (ls.Count > 0)
			//	DrawTable(rtf, ls);
		}

		/// <summary>
		/// Handles emphasis inlines (e.g., *italic*, **bold**, ***bold+italic***, etc.).
		/// We also interpret underscores as underline in this example.
		/// </summary>
		private static void HandleEmphasis(StringBuilder rtf, EmphasisInline emphasisInline, StringBuilder Text = null)
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
			ConvertInline(rtf, emphasisInline, Text);

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
		/// 测量一段文字的宽度，注意需要用等宽字体。
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
				else if (char.GetUnicodeCategory(c).ToString().StartsWith("Other") && c > 256)//:也被算进去了
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
