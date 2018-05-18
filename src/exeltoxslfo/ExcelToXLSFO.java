package exeltoxslfo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * Excelシートから、それらしいXSL-FOファイルを作成します。
 * <pre>
 * Excelでデザインした帳票テンプレートを、それなりの再現度でXSL-FO形式に変換するつもりです。
 * 気に入らないところはXSL-FOをテキストエディタで修正する前提のツールです。
 * </pre>
 */
public class ExcelToXLSFO {

	/**
	 * Logger.
	 */
	private static Logger logger = Logger.getLogger(ExcelToXLSFO.class);
	
	/**
	 * Excel形式の入力ファイルのパス。
	 */
	private String excelFile = null;

	/**
	 * シートインデックス。
	 */
	private int sheetIndex = 0;
	
	/**
	 * XSL-FO:形式の出力ファイルのパス。
	 */
	private String xslFoFile = null;
	
	
	/**
	 * Excelファイルのパスを取得します。
	 * @return Excelファイルのパス。
	 */
	public String getExcelFile() {
		return excelFile;
	}

	/**
	 * Excelファイルのパスを設定します。
	 * @param excelFile Excelファイルのパス。
	 */
	public void setExcelFile(final String excelFile) {
		this.excelFile = excelFile;
	}

	/**
	 * シートインデックスを取得します。
	 * @return シートインデックス。
	 */
	public int getSheetIndex() {
		return sheetIndex;
	}

	/**
	 * シートインデックスを設定します。
	 * @param sheetIndex シートインデックス。
	 */
	public void setSheetIndex(final int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	/**
	 * XSL-FOファイルのパスを取得します。
	 * @return XSL-FOファイルのパス。
	 */
	public String getXslFoFile() {
		return xslFoFile;
	}

	/**
	 * XLS-FOファイルのパスを設定します。
	 * @param xslFoFile XLS-FOファイル。
	 */
	public void setXslFoFile(final String xslFoFile) {
		this.xslFoFile = xslFoFile;
	}

	/**
	 * コマンドラインを解析します。
	 * @param args コマンドライン引数。
	 */
	private void parseAargs(final String[] args) {
		try {
			if (args.length >= 2) {
				for (int i = 0; i < args.length; i++) {
					if ("-s".equals(args[i])) {
						int sheetIndex = Integer.parseInt(args[i + 1]);
						this.setSheetIndex(sheetIndex);
						i++;
					} else {
						if (this.getExcelFile() == null) {
							this.setExcelFile(args[i]);
						} else if (this.getXslFoFile() == null) {
							this.setXslFoFile(args[i]);
						} else {
							throw new Exception();
						}
					}
				}
			} else {
				throw new Exception();
			}
		} catch (NumberFormatException e) {
			logger.error(e.getMessage(), e);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("excel2fo [options] excelfile fofile");
			System.out.println("options:");
			System.out.println("-s sheetidx");
		}
	}
	
	/**
	 * 指定されたWorkbookを取得します。
	 * @return Workbook。
	 * @throws Exception 例外。
	 */
	private Workbook getWorkbook() throws Exception {
		Workbook ret = null;
		FileInputStream is = new FileInputStream(this.getExcelFile());
		try {
			ret = WorkbookFactory.create(is);
		} finally {
			is.close();
		}
		return ret;
	}

	/**
	 * セル情報クラス。
	 *
	 */
	private class CellInfo {
		/**
		 * ワークブック。
		 */
		private Workbook workbook = null;
		/**
		 * セルのスタイル情報。
		 */
		private CellStyle style = null;
		
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
		 * ワークブック。
		 * @param wb ワークブック。
		 */
		public CellInfo(final Workbook wb) {
			this.workbook = wb;
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
		private void setHidden(final boolean hidden) {
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
				logger.debug("birder type=" + this.style.getBorderBottomEnum().name());
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
			this.getBorderStyleAttribute(attrib, "top", this.style.getBorderTopEnum());
			this.getBorderStyleAttribute(attrib, "bottom", this.style.getBorderBottomEnum());
			this.getBorderStyleAttribute(attrib, "left", this.style.getBorderLeftEnum());
			this.getBorderStyleAttribute(attrib, "right", this.style.getBorderRightEnum());
			XSSFCellStyle style = (XSSFCellStyle) this.style;
			this.getBorderColorAttribute(attrib, "border-top-color", style.getTopBorderXSSFColor());
			this.getBorderColorAttribute(attrib, "border-bottom-color", style.getBottomBorderXSSFColor());
			this.getBorderColorAttribute(attrib, "border-left-color", style.getLeftBorderXSSFColor());
			this.getBorderColorAttribute(attrib, "border-right-color", style.getRightBorderXSSFColor());
		}
		
		/**
		 * 背景色のアトリビュートを取得します。
		 * @param attrib 追加する文字列バッファ。
		 */
		public void getBackgroundColorAttribute(final StringBuilder attrib) {
			short colidx = this.style.getFillBackgroundColor();
			logger.debug("background-color=" + colidx);
			XSSFColor c = (XSSFColor) this.style.getFillForegroundColorColor();
			if (c != null) {
				String hexcolor = c.getARGBHex();
				logger.debug("background-color hexcolor=" + hexcolor + ",index=" + c.getIndexed());
				if (hexcolor != null) {
					attrib.append(" background-color=\"#" + hexcolor.substring(2) + "\" ");
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
				logger.debug("fidx=" + fidx);
				Font f = this.workbook.getFontAt((short) fidx);
				attrib.append(" font-family=\"" + f.getFontName() + "\"");
				attrib.append(" font-size=\"" + f.getFontHeightInPoints() + "pt\"");
				XSSFFont xf = (XSSFFont) f;
				XSSFColor color = xf.getXSSFColor();
				String hexcolor = color.getARGBHex();
				logger.debug("hexcolor=" + hexcolor);
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
			if (this.style.getVerticalAlignmentEnum() == VerticalAlignment.TOP) {
				attrib.append(" display-align=\"before\"");
			}
			if (this.style.getVerticalAlignmentEnum() == VerticalAlignment.CENTER) {
				attrib.append(" display-align=\"center\"");
			}
			if (this.style.getVerticalAlignmentEnum() == VerticalAlignment.BOTTOM) {
				attrib.append(" display-align=\"after\"");
			}
			if (this.style.getAlignmentEnum() == HorizontalAlignment.LEFT) {
				attrib.append(" text-align=\"left\"");
			}
			if (this.style.getAlignmentEnum() == HorizontalAlignment.CENTER) {
				attrib.append(" text-align=\"center\"");
			}
			if (this.style.getAlignmentEnum() == HorizontalAlignment.RIGHT) {
				attrib.append(" text-align=\"right\"");
			}
		}

		/**
		 * セルのスタイル情報を設定します。
		 * @param style セルスタイル。
		 */
		public void setStyle(final CellStyle style) {
			this.style = style;
		}
	}

	/**
	 * Excelのテーブル構造を取得します。
	 *
	 */
	private class TableInfo {
		/**
		 * 行の幅リスト。
		 */
		private List<Double> rowHeightList = null;
		/**
		 * カラムの幅リスト。
		 */
		private List<Double> columnWidthList = null;
		
		/**
		 * セル情報。
		 */
		private CellInfo [][] cellInfo = null;
		
		
		/**
		 * 指定されたワークブックのテーブル構造情報を作成します。
		 * @param wb ワークブック。
		 */
		public TableInfo(final Workbook wb) {
			Sheet sh = wb.getSheetAt(getSheetIndex());
			int rows = this.getRows(sh);
			int cols = this.getColums(sh);
			logger.debug("rows=" + rows + ",cols=" + cols);
			this.cellInfo = new CellInfo[rows][cols];
			for (int r = 0; r < rows; r++) {
				for (int c = 0; c < cols; c++) {
					this.cellInfo[r][c] = new CellInfo(wb);
					Cell cell = this.getCell(sh, r, c);
					if (cell != null) {
						this.cellInfo[r][c].setStyle(cell.getCellStyle());
					}
				}
			}
			this.getSpanInfo(wb);
			this.rowHeightList = this.getHeightList(wb, rows);
			this.columnWidthList = this.getWidthList(wb, cols);
			
		}

		/**
		 * 指定されたセルを取得します。
		 * @param sh シート。
		 * @param r 行。
		 * @param c 列。
		 * @return セル。
		 */
		private Cell getCell(final Sheet sh, final int r, final int c) {
			Cell ret = null;
			Row row = sh.getRow(r);
			if (row != null) {
				ret = row.getCell(c);
			}
			return ret;
		}
		
		/**
		 * セル結合情報を取得します。
		 * @param wb ワークブック。
		 */
		private void getSpanInfo(final Workbook wb) {
			Sheet sh = wb.getSheetAt(getSheetIndex());
			int n = sh.getNumMergedRegions();
			for (int i = 0; i < n; i++) {
				logger.debug("mergedRegion=" + i);
				CellRangeAddress rgn = sh.getMergedRegion(i);
				int r0 = rgn.getFirstRow();
				int c0 = rgn.getFirstColumn();
				int rowSpan = rgn.getLastRow() - rgn.getFirstRow() + 1;
				int colSpan = rgn.getLastColumn() - rgn.getFirstColumn() + 1;
				logger.debug("cell(" + r0 + "," + c0 + ") rowspan=" + rowSpan + ",colspan=" + colSpan);
				for (int r = r0; r <= rgn.getLastRow(); r++) {
					for (int c = c0; c <= rgn.getLastColumn(); c++) {
						this.getCellInfo(r, c).setHidden(true);
						logger.debug("cell(" + r + "," + c + ") is hidden.");
					}
				}
				this.getCellInfo(r0, c0).setRowSpan(rowSpan);
				this.getCellInfo(r0, c0).setColumnSpan(colSpan);
				this.getCellInfo(r0, c0).setHidden(false);
			}
		}
		
		/**
		 * セル情報を取得します。
		 * @param row 行。
		 * @param col カラム。
		 * @return セル情報。
		 */
		public CellInfo getCellInfo(final int row, final int col) {
			return this.cellInfo[row][col];
		}

		/**
		 * Excelシート中のテーブル行数を取得します。
		 * @param sh シート。
		 * @return 行数。
		 */
		private int getRows(final Sheet sh) {
			int rows = sh.getLastRowNum() + 1;
			return rows;
		}

		/**
		 * Excelシート中のテーブルカラム数を取得します。
		 * @param sh シート。
		 * @return 行数。
		 */
		private int getColums(final Sheet sh) {
			int cols = 0;
			logger.debug("first=" + sh.getFirstRowNum() + ",last=" + sh.getLastRowNum());
			for (int i = sh.getFirstRowNum(); i <= sh.getLastRowNum(); i++) {
				Row r = sh.getRow(i);
				if (r != null) {
					logger.debug("r.getLastCellNum()=" + r.getLastCellNum());
					if (cols < r.getLastCellNum()) {
						cols = r.getLastCellNum();
					}
				}
			}
			return cols;
		}
		
		/**
		 * Excelのテーブル行の高さの配列を取得します。
		 * @param wb ワークブック。
		 * @param rows テーブル行数。
		 * @return テーブル行の高さ(単位ポイント)の配列を取得します。
		 */
		private List<Double> getHeightList(final Workbook wb, final int rows) {
			Sheet sh = wb.getSheetAt(getSheetIndex());
			List<Double> ret = new ArrayList<Double>();
			for (int i = 0; i < rows; i++) {
				Row r = sh.getRow(i);
				if (r != null) {
					double h = r.getHeightInPoints();
					logger.debug("h=" + h);
					ret.add(h);
				} else {
					double h = sh.getDefaultRowHeightInPoints();
					logger.debug("h=" + h);
					ret.add(h);
				}
			}
			return ret;
		}
		
		/**
		 * Excelのテーブルカラムの幅の配列を取得します。
		 * <pre>
		 * カラム幅のポイントへの変換をポイントに正確に変換するのは困難なようです。
		 * </pre>
		 * @param wb ワークブック。
		 * @param cols テーブルのカラム数。
		 * @return テーブルカラムの幅(単位ポイント)の配列を取得します。
		 */
		private List<Double> getWidthList(final Workbook wb, final int cols) {
			Font f = wb.getFontAt((short) 0);
			logger.debug("points=" + f.getFontHeightInPoints());
			Sheet sh = wb.getSheetAt(getSheetIndex());
			List<Double> ret = new ArrayList<Double>();
			for (int i = 0; i < cols; i++) {
				// セル幅の計算はかなり適当
				double w = sh.getColumnWidth(i) / 256.0 * (f.getFontHeightInPoints() * 0.56);
				logger.debug("w=" + w);
				ret.add(w);
			}
			return ret;
		}
		

		/**
		 * 行の高さリストを取得します。
		 * @return 行の高さリスト。
		 */
		public List<Double> getRowHeightList() {
			return rowHeightList;
		}

		/**
		 * カラムの幅リストを取得します。
		 * @return カラムの幅リスト。
		 */
		public List<Double> getColumnWidthList() {
			return columnWidthList;
		}
		
		/**
		 * テーブルの行数を取得します。
		 * @return テーブルの行数。
		 */
		public int getRows() {
			return this.rowHeightList.size();
		}
		
		/**
		 * テーブルのカラム数を取得します。
		 * @return テーブルのカラム数。
		 */
		public int getColumns() {
			return this.columnWidthList.size();
		}
		
		/**
		 * テーブル幅を取得します。
		 * @return テーブル幅。
		 */
		public double getTableWidth() {
			double ret = 0.0;
			for (Double w: this.columnWidthList) {
				ret += w.doubleValue();
			}
			return ret;
		}
		

		/**
		 * 指定された行のアトリビュートを取得します。
		 * @param r 行インデックス。
		 * @return アトリビュート文字列。
		 */
		public String getRowAttribute(final int r) {
			double h = this.getRowHeightList().get(r);
			String attrib = "height=\"" + h + "pt\"";
			return attrib;
		}
	}

	/**
	 * ワークブックのテーブル情報を取得します。
	 * @param wb ワークブック。
	 * @return テーブル情報。
	 */
	private TableInfo getTableInfo(final Workbook wb) {
		TableInfo ret = new TableInfo(wb);
		return ret;
	}
	
	/**
	 * XMLのルート開始タグ。
	 */
	private static final String XML_ROOT_BEGIN = 
			"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + 
			"<fo:root xmlns:fo=\"http://www.w3.org/1999/XSL/Format\" xml:lang=\"ja\">\n";
	/**
	 * XMLのメート終了タグ。
	 */
	private static final String XML_ROOT_END = "</fo:root>\n";
	
	/**
	 * A4縦用のページマスタ。
	 */
	private static final String PAGE_MASTER_A4 =
			"	<fo:layout-master-set>\n" + 
			"		<fo:simple-page-master page-height=\"297mm\" page-width=\"210mm\" margin-top=\"10mm\" margin-left=\"20mm\" margin-right=\"20mm\" margin-bottom=\"10mm\" master-name=\"PageMaster\">\n" + 
			"			<fo:region-body margin-top=\"20mm\" margin-left=\"0mm\" margin-right=\"0mm\" margin-bottom=\"10mm\"/>\n" + 
			"		</fo:simple-page-master>\n" + 
			"	</fo:layout-master-set>\n"; 
	
	/**
	 * ページの開始タグ。
	 */
	private static final String PAGE_BEGIN = 
			"	<fo:page-sequence initial-page-number=\"1\" master-reference=\"PageMaster\" font-family=\"${fontName}\" font-size=\"${fontPoint}pt\">\n" + 
			"		<fo:flow flow-name=\"xsl-region-body\">\n" + 
			"			<fo:block  space-before=\"1em\" >\n";
	
	/**
	 * ページ終了タグ。
	 */
	private static final String PAGE_END = 
			"			</fo:block>\n" + 
			"		</fo:flow>\n" + 
			"	</fo:page-sequence>\n"; 
	
	/**
	 * テーブル開始タグ。
	 */
	private static final String TABLE_BEGIN = 
			"				<fo:table inline-progression-dimension=\"${width}pt\" table-layout=\"fixed\">\n"; 
	
	/**
	 * テーブル終了タグ。
	 */
	private static final String TABLE_END = 
			"				</fo:table>\n"; 
	
	/**
	 * カラム幅設定タグ。
	 */
	private static final String COLUMN_WIDTH = 
			"					<fo:table-column column-number=\"${cidx}\" column-width=\"${width}pt\" />\n"; 
	
	/**
	 * テーブルボディ開始タグ。
	 */
	private static final String TABLE_BODY_BEGIN = 
			"					<fo:table-body>\n"; 
	
	/**
	 * テーブルボディ終了タグ。
	 */
	private static final String TABLE_BODY_END = 
			"					</fo:table-body>\n";
	
	/**
	 * テーブル行開始タグ。
	 */
	private static final String TABLE_ROW_BEGIN = 
			"						<fo:table-row ${attrib}>\n"; 
	
	/**
	 * テーブル行終了タグ。
	 */
	private static final String TABLE_ROW_END = 
			"						</fo:table-row>\n"; 
	
	/**
	 * セル開始タグ。
	 */
	private static final String TABLE_CELL_BEGIN = 
			"							<fo:table-cell ${attrib}>\n"; 
	
	/**
	 * セル終了タグ。
	 */
	private static final String TABLE_CELL_END = 
			"							</fo:table-cell>\n"; 

	/**
	 * セル内容ブロックタグ。
	 */
	private static final String TABLE_CELL_BLOCK_BEGIN = 
			"								<fo:block>";

	
	/**
	 * セル内容ブロックタグ。
	 */
	private static final String TABLE_CELL_BLOCK_END = 
			"</fo:block>\n";
	
	/**
	 * セルの値を取得します。
	 * @param cell セル。
	 * @return 値。
	 */
	private String getCellValue(final Cell cell) {
		String value = "";
		if (cell.getCellTypeEnum() == CellType.STRING) {
			value = cell.getStringCellValue();
		} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
			double v  = cell.getNumericCellValue();
			value = "" + v;
		}
		return value;
	}
	
	/**
	 * 指定された行のテーブルセルのXMLを作成します。
	 * 
	 * @param wb ワークブック。
	 * @param tinfo テーブル情報。
	 * @param r 行インデックス。
	 * @return XMLの文字列。
	 */
	private String getTableCellsXml(final Workbook wb, final TableInfo tinfo, final int r) {
		Sheet sh = wb.getSheetAt(getSheetIndex());
		StringBuilder sb = new StringBuilder();
		Row row = sh.getRow(r);
		if (row != null) {
			for (int c = 0; c < tinfo.getColumns(); c++) {
				CellInfo ci = tinfo.getCellInfo(r, c);
				if (ci.isHidden()) {
					continue;
				}
				Cell cell = row.getCell(c);
				if (cell != null) {
					String cellBegin = TABLE_CELL_BEGIN.replaceAll("\\$\\{attrib\\}", ci.getCellAttribute());
					sb.append(cellBegin);
					String value = this.getCellValue(cell);
					sb.append(TABLE_CELL_BLOCK_BEGIN);
					sb.append(value);
					sb.append(TABLE_CELL_BLOCK_END);
					sb.append(TABLE_CELL_END);
				} else {
					String attrib = "";
					String cellBegin = TABLE_CELL_BEGIN.replaceAll("\\$\\{attrib\\}", attrib);
					sb.append(cellBegin);
					sb.append(TABLE_CELL_BLOCK_BEGIN);
					sb.append(TABLE_CELL_BLOCK_END);
					sb.append(TABLE_CELL_END);
				}
			}
		} else {
			for (int c = 0; c < tinfo.getColumns(); c++) {
				CellInfo ci = tinfo.getCellInfo(r, c);
				if (ci.isHidden()) {
					continue;
				}
				String attrib = "";
				String cellBegin = TABLE_CELL_BEGIN.replaceAll("\\$\\{attrib\\}", attrib);
				sb.append(cellBegin);
				sb.append(TABLE_CELL_BLOCK_BEGIN);
				sb.append(TABLE_CELL_BLOCK_END);
				sb.append(TABLE_CELL_END);
			}
		}
		return sb.toString();
	}
	
	/**
	 * テーブルのXMLを作成します。
	 * @param wb ワークブック。
	 * @param tinfo テーブル情報。
	 * @return XML文字列。
	 */
	private String getTableXml(final Workbook wb, final TableInfo tinfo) {
		StringBuilder sb = new StringBuilder();
		String tblbegin = TABLE_BEGIN.replaceAll("\\$\\{width\\}", "" + tinfo.getTableWidth()); 
		sb.append(tblbegin);
		for (int i = 0; i < tinfo.getColumns(); i++) {
			String colinfo = COLUMN_WIDTH.replaceAll("\\$\\{width\\}", "" + tinfo.getColumnWidthList().get(i).doubleValue());
			colinfo = colinfo.replaceAll("\\$\\{cidx\\}", "" + (i + 1));
			sb.append(colinfo);
		}
		sb.append(TABLE_BODY_BEGIN);
		for (int r = 0; r < tinfo.getRows(); r++) {
			String attrib = tinfo.getRowAttribute(r);
			String tableRowBegin = TABLE_ROW_BEGIN.replaceAll("\\$\\{attrib\\}", attrib);
			sb.append(tableRowBegin);
			sb.append(this.getTableCellsXml(wb, tinfo, r));
			sb.append(TABLE_ROW_END);
		}
		sb.append(TABLE_BODY_END);
		sb.append(TABLE_END);
		return sb.toString();
	}
	
	/**
	 * XSL-FO形式のXMLを取得します。
	 * @param wb ワークブック。
	 * @param tinfo テーブル情報。
	 * @return XSL-FO
	 */
	private String getXSLFO(final Workbook wb, final TableInfo tinfo) {
		StringBuilder sb = new StringBuilder();
		sb.append(XML_ROOT_BEGIN);
		sb.append(PAGE_MASTER_A4);
		Font f = wb.getFontAt((short) 0);
		String pageBegin = PAGE_BEGIN;
		pageBegin = pageBegin.replaceAll("\\$\\{fontName\\}", f.getFontName());
		pageBegin = pageBegin.replaceAll("\\$\\{fontPoint\\}", "" + f.getFontHeightInPoints());
		sb.append(pageBegin);
		sb.append(this.getTableXml(wb, tinfo));
		sb.append(PAGE_END);
		sb.append(XML_ROOT_END);
		return sb.toString();
	}
	
	/**
	 * ExcelファイルからXSL-FO形式のXMLを作成します。
	 * @return XSL-FO形式の文字列。
	 * @throws Exception 例外。
	 */
	public String convert() throws Exception {
		Workbook wb = this.getWorkbook();
		logger.debug("wb sheet count=" + wb.getNumberOfSheets());
		TableInfo tinfo = this.getTableInfo(wb);
		String xml = this.getXSLFO(wb, tinfo);
		logger.debug("so=" + xml);
		if (this.xslFoFile != null) {
			FileOutputStream os = new FileOutputStream(this.xslFoFile);
			try {
				os.write(xml.getBytes("utf-8"));
			} finally {
				os.close();
			}
		}
		return xml;
	}
	
	/**
	 * メイン処理。
	 * 
	 * @param args コマンドライン引数。
	 */
	public static void main(final String[] args) {
		ExcelToXLSFO conv = new ExcelToXLSFO();
		conv.parseAargs(args);
		try {
			conv.convert();
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
}
