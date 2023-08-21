package jp.dataforms.exeltoxslfo;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * セル情報クラス。
 *
 */
public class CellInfo {

	/**
	 * Logger.
	 */
	private Logger logger = LogManager.getLogger(CellInfo.class);

	/**
	 * ワークブック。
	 */
	private Workbook workbook = null;

	/**
	 * セルのスタイル情報。
	 */
	private CellStyle style = null;


	/**
	 * セルのスタイル情報。
	 */
	private CellStyle bottomRightStyle = null;


	/**
	 * 説結合によって表示されないセルを示すフラグ。
	 */
	private boolean hidden = false;

	/**
	 * Row spanの値。
	 */
	private int rowSpan = -1;

	/**
	 * Column spanの値。
	 */
	private int columnSpan = -1;

	/**
	 * セルの値。
	 */
	private String value = null;

	/**
	 * セルタイプ。
	 */
	private CellType cellType = null;

	/**
	 * 行。
	 */
	private int row = 0;

	/**
	 * 列。
	 */
	private int column = 0;

	/**
	 * ワークブック。
	 * @param wb ワークブック。
	 * @param row 行。
	 * @param col 列。
	 */
	public CellInfo(final Workbook wb, final int row, final int col) {
		this.workbook = wb;
		this.row = row;
		this.column = col;
	}

	/**
	 * Column spanの値を取得します。
	 * @return Column spanの値
	 */
	public int getColumnSpan() {
		return columnSpan;
	}

	/**
	 * Column spanの値を設定します。
	 * @param columnSpan Column spanの値。
	 */
	public void setColumnSpan(final int columnSpan) {
		this.columnSpan = columnSpan;
	}

	/**
	 * Row spanの値を取得します。
	 * @return Row spanの値。
	 */
	public int getRowSpan() {
		return rowSpan;
	}

	/**
	 * Row spanの値を設定します。
	 * @param rowSpan Row spanの値。
	 */
	public void setRowSpan(final int rowSpan) {
		this.rowSpan = rowSpan;
	}

	/**
	 * セル結合によって非表示になったセルを判定します。
	 * @return 非表示セルの場合true。
	 */
	public boolean isHidden() {
		return hidden;
	}

	/**
	 * 非表示セルであることを設定します。
	 * @param hidden 非表示セルの場合true。
	 */
	public void setHidden(final boolean hidden) {
		this.hidden = hidden;
	}

	/**
	 * セルのアトリビュートを取得します。
	 * @return セルのアトリビュート文字列。
	 */
	public String getCellAttribute() {
		StringBuilder attrib = new StringBuilder();
		if (this.getRowSpan() > 1) {
			attrib.append(" number-rows-spanned=\"" + this.getRowSpan() + "\" ");
		}
		if (this.getColumnSpan() >= 0) {
			attrib.append(" number-columns-spanned=\"" + this.getColumnSpan() + "\" ");
		}
		if (this.style != null) {
			this.getAlignmentAttribute(attrib);
			this.getFontAttribute(attrib);
			this.getBackgroundColorAttribute(attrib);
			this.getBorderAttribute(attrib);
		}
		return attrib.toString();
	}

	/**
	 * ExcelのBorderStyleをXSL-FOのborder-styleに変換します。
	 * @param style ExcelのBorderStyle。
	 * @return XSL-FOのborder-style。
	 */
	private String getBorderStyle(final BorderStyle style) {
		String ret = null;
		if (style == BorderStyle.HAIR) {
			ret = "dotted";
		} else if (style == BorderStyle.DOTTED) {
			ret = "dotted";
		} else if (style == BorderStyle.DASH_DOT_DOT) {
			ret = "dashed";
		} else if (style == BorderStyle.DASH_DOT) {
			ret = "dashed";
		} else if (style == BorderStyle.DASHED) {
			ret = "dashed";
		} else if (style == BorderStyle.THIN) {
			ret = "solid";
		} else if (style == BorderStyle.MEDIUM_DASH_DOT_DOT) {
			ret = "dashed";
		} else if (style == BorderStyle.SLANTED_DASH_DOT) {
			ret = "dashed";
		} else if (style == BorderStyle.MEDIUM_DASH_DOT) {
			ret = "dashed";
		} else if (style == BorderStyle.MEDIUM_DASHED) {
			ret = "dashed";
		} else if (style == BorderStyle.MEDIUM) {
			ret = "solid";
		} else if (style == BorderStyle.THICK) {
			ret = "solid";
		} else if (style == BorderStyle.DOUBLE) {
			ret = "double";
		}
		return ret;
	}

	/**
	 * ExcelのBorderStyleをXSL-FOのborder-widthに変換します。
	 * @param style ExcelのBorderStyle。
	 * @return XSL-FOのborder-width。
	 */
	private String getBorderWidth(final BorderStyle style) {
		String ret = null;
		if (style == BorderStyle.HAIR) {
			ret = "0.12mm";
		} else if (style == BorderStyle.DOTTED) {
			ret = "thin";
		} else if (style == BorderStyle.DASH_DOT_DOT) {
			ret = "thin";
		} else if (style == BorderStyle.DASH_DOT) {
			ret = "thin";
		} else if (style == BorderStyle.DASHED) {
			ret = "thin";
		} else if (style == BorderStyle.THIN) {
			ret = "thin";
		} else if (style == BorderStyle.MEDIUM_DASH_DOT_DOT) {
			ret = "medium";
		} else if (style == BorderStyle.SLANTED_DASH_DOT) {
			ret = "medium";
		} else if (style == BorderStyle.MEDIUM_DASH_DOT) {
			ret = "medium";
		} else if (style == BorderStyle.MEDIUM_DASHED) {
			ret = "medium";
		} else if (style == BorderStyle.MEDIUM) {
			ret = "medium";
		} else if (style == BorderStyle.THICK) {
			ret = "thick";
		} else if (style == BorderStyle.DOUBLE) {
			ret = "1.2mm";
		}
		return ret;
	}


	/**
	 * ボーダースタイルのアトリビュートを作成します。
	 * @param attrib アトリビュートを追加する文字列バッファ。
	 * @param prop top,bottom,left,rightのいずれかを指定。
	 * @param style BorderStyle。
	 */
	private void getBorderStyleAttribute(final StringBuilder attrib, final String prop, final BorderStyle style) {
		if (style != BorderStyle.NONE) {
			attrib.append(" border-" + prop + "-style=\"" + this.getBorderStyle(style) +"\"");
			attrib.append(" border-" + prop + "-width=\"" + this.getBorderWidth(style) +"\"");
		}
	}

	/**
	 * ボーダーの色アトリビュートを作成します。
	 * @param attrib アトリビュートを追加する文字列バッファ。
	 * @param prop top,bottom,left,rightのいずれかを指定。
	 * @param color ボーダーの色。
	 */
	private void getBorderColorAttribute(final StringBuilder attrib, final String prop, final XSSFColor color) {
		if (color != null) {
			String cc = color.getARGBHex();
			if (cc != null) {
				attrib.append(" " + prop + "=\"#" + cc.substring(2) +"\"");
			}
		}
	}

	/**
	 * Border関連のアトリビュートを作成します。
	 * @param attrib アトリビュートを追加する文字列バッファ。
	 */
	private void getBorderAttribute(final StringBuilder attrib) {
		this.getBorderStyleAttribute(attrib, "top", this.style.getBorderTop());
		this.getBorderStyleAttribute(attrib, "left", this.style.getBorderLeft());
		if (this.bottomRightStyle == null) {
			this.getBorderStyleAttribute(attrib, "bottom", this.style.getBorderBottom());
			this.getBorderStyleAttribute(attrib, "right", this.style.getBorderRight());
		} else {
			this.getBorderStyleAttribute(attrib, "bottom", this.bottomRightStyle.getBorderBottom());
			this.getBorderStyleAttribute(attrib, "right", this.bottomRightStyle.getBorderRight());
		}
		XSSFCellStyle style = (XSSFCellStyle) this.style;
		this.getBorderColorAttribute(attrib, "border-top-color", style.getTopBorderXSSFColor());
		this.getBorderColorAttribute(attrib, "border-left-color", style.getLeftBorderXSSFColor());
		if (this.bottomRightStyle != null) {
			style = (XSSFCellStyle) this.bottomRightStyle;
		}
		this.getBorderColorAttribute(attrib, "border-bottom-color", style.getBottomBorderXSSFColor());
		this.getBorderColorAttribute(attrib, "border-right-color", style.getRightBorderXSSFColor());
	}

	/**
	 * 背景色のアトリビュートを取得します。
	 * @param attrib 追加する文字列バッファ。
	 */
	public void getBackgroundColorAttribute(final StringBuilder attrib) {
		XSSFColor c = (XSSFColor) this.style.getFillForegroundColorColor();
		int cidx = this.style.getFillForegroundColor();
		logger.debug("cidx=" + cidx);
		if (c != null) {
			byte[] rgb = c.getRGBWithTint();
			String hexcolor = String.format("%02x", rgb[0]) + String.format("%02x", rgb[1]) + String.format("%02x", rgb[2]);
			if (hexcolor != null) {
				logger.debug("row,col=(" + this.row + "," + this.column + "), hexcolor=" + hexcolor + ", cidx=" + cidx);
				attrib.append(" background-color=\"#" + hexcolor + "\" ");
			}
		}
	}

	/**
	 * フォント関連情報を取得します。
	 * @param attrib 追加する文字列バッファ。
	 */
	public void getFontAttribute(final StringBuilder attrib) {
		int fidx = this.style.getFontIndex();
		if (fidx > 0) {
			Font f = this.workbook.getFontAt(fidx);
			attrib.append(" font-family=\"" + f.getFontName() + "\"");
			attrib.append(" font-size=\"" + f.getFontHeightInPoints() + "pt\"");
			XSSFFont xf = (XSSFFont) f;
			XSSFColor color = xf.getXSSFColor();
			String hexcolor = color.getARGBHex();
			attrib.append(" color=\"#" + hexcolor.substring(2) + "\"");
			if (f.getBold()) {
				attrib.append(" font-weight=\"bold\"");
			}
			if (f.getItalic()) {
				attrib.append(" font-style=\"italic\"");
			}
			byte u = f.getUnderline();
			if (u == 1) {
				attrib.append(" text-decoration=\"underline\"");
			}
		}
	}

	/**
	 * 配置情報の属性を追加します。
	 * @param attrib 追加する文字列バッファ。
	 */
	public void getAlignmentAttribute(final StringBuilder attrib) {
		if (this.style.getVerticalAlignment() == VerticalAlignment.TOP) {
			attrib.append(" display-align=\"before\"");
		}
		if (this.style.getVerticalAlignment() == VerticalAlignment.CENTER) {
			attrib.append(" display-align=\"center\"");
		}
		if (this.style.getVerticalAlignment() == VerticalAlignment.BOTTOM) {
			attrib.append(" display-align=\"after\"");
		}
		if (this.style.getAlignment() == HorizontalAlignment.LEFT) {
			attrib.append(" text-align=\"left\"");
		} else 	if (this.style.getAlignment() == HorizontalAlignment.CENTER) {
			attrib.append(" text-align=\"center\"");
		} else if (this.style.getAlignment() == HorizontalAlignment.RIGHT) {
			attrib.append(" text-align=\"right\"");
		} else {
			CellType type = this.getCellType();
			if (type == CellType.NUMERIC) {
				attrib.append(" text-align=\"right\"");
			}
		}
	}

	/**
	 * セルのスタイル情報を設定します。
	 * @param style セルスタイル。
	 */
	public void setStyle(final CellStyle style) {
		this.style = style;
	}

	/**
	 * セルの右下のスタイルを設定します。
	 * <pre>
	 * セルが結合された場合のみ設定。
	 * </pre>
	 * @param bottomRightStyle セルの右下のスタイル。
	 */
	public void setBottomRightStyle(final CellStyle bottomRightStyle) {
		this.bottomRightStyle = bottomRightStyle;
	}

	/**
	 * セルの値を取得します。
	 * @return セルの値。
	 */
	public String getValue() {
		return value;
	}

	/**
	 * セルの値を設定します。
	 * @param value セルの値。
	 */
	public void setValue(final String value) {
		this.value = value;
	}

	/**
	 * セルタイプを取得します。
	 * @return セルタイプ。
	 */
	public CellType getCellType() {
		return cellType;
	}

	/**
	 * セルタイプを設定します。
	 * @param cellType セルタイプ。
	 */
	public void setCellType(final CellType cellType) {
		this.cellType = cellType;
	}
}
