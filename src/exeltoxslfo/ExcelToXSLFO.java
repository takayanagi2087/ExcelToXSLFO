package exeltoxslfo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;

import net.arnx.jsonic.JSON;

/**
 * Excelシートから、それらしいXSL-FOファイルを作成します。
 * <pre>
 * Excelでデザインした帳票テンプレートを、それなりの再現度でXSL-FO形式に変換するつもりです。
 * 気に入らないところはXSL-FOをテキストエディタで修正する前提のツールです。
 * </pre>
 */
public class ExcelToXSLFO {

	/**
	 * Logger.
	 */
	private static Logger logger = LogManager.getLogger(ExcelToXSLFO.class);

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
	 * イメージフィールドに対応した画像タグ。
	 */
	private StringBuilder vImageList = null;

	/**
	 * コンストラクタ。
	 */
	public ExcelToXSLFO() {
		this.vImageList = new StringBuilder();
	}

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
	 * 引数指定の例外。
	 *
	 */
	private class ArgException extends Exception {

	}

	/**
	 * コマンドラインを解析します。
	 * @param args コマンドライン引数。
	 * @throws Exception 例外。
	 */
	private void parseAargs(final String[] args) throws Exception {
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
						throw new ArgException();
					}
				}
			}
			if (this.getExcelFile() == null || this.getXslFoFile() == null) {
				throw new ArgException();
			}
		} else {
			throw new ArgException();
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
		 * 画像情報。
		 */
		private List<ImageInfo> imageList = new ArrayList<ImageInfo>();

		/**
		 * 指定されたワークブックのテーブル構造情報を作成します。
		 * @param wb ワークブック。
		 * @throws Exception 例外。
		 */
		public TableInfo(final Workbook wb) throws Exception {
			Sheet sh = wb.getSheetAt(getSheetIndex());
			FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
			int rows = this.getRows(sh);
			int cols = this.getColums(sh) + 1;
			this.cellInfo = new CellInfo[rows][cols];
			for (int r = 0; r < rows; r++) {
				for (int c = 0; c < cols; c++) {
					this.cellInfo[r][c] = new CellInfo(wb, r, c);
					Cell cell = this.getCell(sh, r, c);
					if (cell != null) {
						this.cellInfo[r][c].setStyle(cell.getCellStyle());
						this.cellInfo[r][c].setValue(ExcelToXSLFO.this.getCellValue(cell, fe));
						if (cell.getCellType() == CellType.FORMULA) {
							CellValue cv = fe.evaluate(cell);
							this.cellInfo[r][c].setCellType(cv.getCellType());
						} else {
							this.cellInfo[r][c].setCellType(cell.getCellType());
						}
					}
				}
			}
			this.getSpanInfo(wb);
			this.rowHeightList = this.getHeightList(wb, rows);
			this.columnWidthList = this.getWidthList(wb, cols);
			XSSFDrawing drawing = (XSSFDrawing) sh.createDrawingPatriarch();
			List<XSSFShape> shapeList = drawing.getShapes();
			for (XSSFShape shape: shapeList) {
				if (shape instanceof XSSFPicture) {
					XSSFPicture pic = (XSSFPicture) shape;
					XSSFClientAnchor  anc = (XSSFClientAnchor) pic.getAnchor();
					double top = this.getTop(anc.getRow1()) + anc.getDy1() / Units.EMU_PER_POINT;
					double left = this.getLeft(anc.getCol1()) + anc.getDx1() / Units.EMU_PER_POINT;
					double bottom = this.getTop(anc.getRow2()) + anc.getDy2() / Units.EMU_PER_POINT;
					double right = this.getLeft(anc.getCol2()) + anc.getDx2() / Units.EMU_PER_POINT;
					double height = bottom - top + 1;
					double width = right - left + 1;
					this.imageList.add(new ImageInfo(top, left, height, width, pic.getPictureData()));
				}
			}
		}


		/**
		 * 画像リストを取得します。
		 * @return 画像リスト。
		 */
		public List<ImageInfo> getImageList() {
			return imageList;
		}

		/**
		 * 指定した行の上端座標(pt)を取得します。
		 * @param row 行インデックス。
		 * @return 上端座標(pt)。
		 */
		private double getTop(final int row) {
			double ret = 0;
			for (int i = 0; i < row; i++) {
				Double h = this.rowHeightList.get(i);
				ret += h;
			}
			return ret;
		}

		/**
		 * 指定したセルの左端座標(pt)を取得します。
		 * @param cell セルインデックス。
		 * @return セルの左端座標(pt)。
		 */
		private double getLeft(final int cell) {
			double ret = 0;
			for (int i = 0; i < cell; i++) {
				Double h = this.columnWidthList.get(i);
				ret += h;
			}
			return ret;
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
				CellRangeAddress rgn = sh.getMergedRegion(i);
				int r0 = rgn.getFirstRow();
				int c0 = rgn.getFirstColumn();
				int rowSpan = rgn.getLastRow() - rgn.getFirstRow() + 1;
				int colSpan = rgn.getLastColumn() - rgn.getFirstColumn() + 1;
				for (int r = r0; r <= rgn.getLastRow(); r++) {
					for (int c = c0; c <= rgn.getLastColumn(); c++) {
						this.getCellInfo(r, c).setHidden(true);
					}
				}
				this.getCellInfo(r0, c0).setRowSpan(rowSpan);
				this.getCellInfo(r0, c0).setColumnSpan(colSpan);
				this.getCellInfo(r0, c0).setHidden(false);
				Row row = sh.getRow(rgn.getLastRow());
				if (row != null) {
					Cell cell = row.getCell(rgn.getLastColumn());
					if (cell != null) {
						this.getCellInfo(r0, c0).setBottomRightStyle(cell.getCellStyle());
					}
				}
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
			for (int i = sh.getFirstRowNum(); i <= sh.getLastRowNum(); i++) {
				Row r = sh.getRow(i);
				if (r != null) {
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
					ret.add(h);
				} else {
					double h = sh.getDefaultRowHeightInPoints();
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
			Font f = wb.getFontAt(0);
			Sheet sh = wb.getSheetAt(getSheetIndex());
			List<Double> ret = new ArrayList<Double>();
			for (int i = 0; i < cols; i++) {
				// セル幅の計算はかなり適当
				double w = sh.getColumnWidth(i) / 256.0 * (f.getFontHeightInPoints() * 0.56);
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
	 * @throws Exception 例外。
	 */
	private TableInfo getTableInfo(final Workbook wb) throws Exception {
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
			"								<fo:block margin-left=\"1mm\">";


	/**
	 * セル内容ブロックタグ。
	 */
	private static final String TABLE_CELL_BLOCK_END =
			"</fo:block>\n";


	/**
	 * セルの値を取得します。
	 * @param cell セル。
	 * @param fe 数式評価ツール。
	 * @return 値。
	 */
	private String getCellValue(final Cell cell, final FormulaEvaluator fe) {
		DataFormatter fmt = new DataFormatter();
		String value = "";
		if (cell.getCellType() == CellType.BLANK) {
			value = "";
		} else if (cell.getCellType() == CellType.STRING) {
			value = cell.getStringCellValue();
		} else if (cell.getCellType() == CellType.FORMULA) {
			value = fmt.formatCellValue(cell, fe);
		} else {
			value = fmt.formatCellValue(cell);
		}
		return value;
	}

	/**
	 * Jsonから変換したMapからの値取得。
	 * @param info Jsonから変換したMap。
	 * @param key キー。
	 * @param dv デフォルト値、
	 * @return 値。
	 */
	protected BigDecimal getBigDecimalValue(final Map<String, Object> info, final String key, final BigDecimal dv) {
		BigDecimal value = (BigDecimal) info.get(key);
		if (value == null) {
			value = dv;
		}
		return value;

	}

	/**
	 * 画像タグを取得します。
	 * @param tinfo テーブル情報。
	 * @param cell セル。
	 * @param ci セル情報。
	 * @param tag タグ。
	 * @param json 画像パラメータのJson、
	 * @return 画像タグ。
	 */
	protected String getImageTag(final TableInfo tinfo, final Cell cell, final CellInfo ci, final String tag, final String json) {
		@SuppressWarnings("unchecked")
		Map<String, Object> info = (Map<String, Object>) JSON.decode(json, HashMap.class);
		int r0 = cell.getRowIndex();
		int c0 = cell.getColumnIndex();
		BigDecimal rows = this.getBigDecimalValue(info, "rows", BigDecimal.valueOf(1));
		BigDecimal cols = this.getBigDecimalValue(info, "columns", BigDecimal.valueOf(1));
		int r1 = r0 + rows.intValue();
		int c1 = c0 + cols.intValue();

		BigDecimal dx1 = this.getBigDecimalValue(info, "dx1", BigDecimal.valueOf(0));
		BigDecimal dy1 = this.getBigDecimalValue(info, "dy1", BigDecimal.valueOf(0));
		BigDecimal dx2 = this.getBigDecimalValue(info, "dx2", BigDecimal.valueOf(0));
		BigDecimal dy2 = this.getBigDecimalValue(info, "dy2", BigDecimal.valueOf(0));

		double top = tinfo.getTop(r0) + dy1.intValue();
		double left = tinfo.getLeft(c0) + dx1.intValue();
		double bottom = tinfo.getTop(r1) + dy2.intValue();
		double right = tinfo.getLeft(c1) + dx2.intValue();
		double height = bottom - top + 1;
		double width = right - left + 1;

		StringBuilder sb = new StringBuilder();
		String imageBlockBegin = IMAGE_BLOCK_BEGIN;
		imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{top\\}", "" + top);
		imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{left\\}", "" + left);
		imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{height\\}", "" + height);
		imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{width\\}", "" + width);
		String aspect = (String) info.get("aspect");
		String scaling = "non-uniform";
		if ("image".equals(aspect)) {
			scaling = "uniform";
		}
		sb.append(imageBlockBegin);
		sb.append("					<fo:block><fo:external-graphic src=\"" + tag + "\" width=\"" + width + "pt\" height=\"" + height + "pt\" content-width=\"" + width + "pt\" content-height=\"" + height + "pt\" border-style=\"dotted\" border-width=\"0mm\" scaling=\"" + scaling + "\"/></fo:block>\n");
		sb.append(IMAGE_BLOCK_END);

		return sb.toString();
	}

	/**
	 * セルの値を取得します。
	 * <pre>
	 * セルに画像用のタグがあった場合、画像に展開します。
	 * </pre>
	 * @param tinfo テーブル情報。
	 * @param cell セル。
	 * @param ci セル情報。
	 * @return セルの値。
	 */
	protected String getCellValue(final TableInfo tinfo, final Cell cell, final CellInfo ci) {
		Pattern p = Pattern.compile("(\\$\\{.+?\\})(\\{.+?\\})");
		Matcher m = p.matcher(ci.getValue());
		if (m.find()) {
			vImageList.append(this.getImageTag(tinfo, cell, ci, m.group(1), m.group(2)));
			return "";
		} else {
			return ci.getValue();
		}
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
					String value = this.getCellValue(tinfo, cell, ci);
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
		String ret = sb.toString();
		return ret;
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
			String cells = this.getTableCellsXml(wb, tinfo, r);
			String tableRowBegin = TABLE_ROW_BEGIN.replaceAll("\\$\\{attrib\\}", attrib);
			sb.append(tableRowBegin);
			sb.append(cells);
			sb.append(TABLE_ROW_END);
		}
		sb.append(TABLE_BODY_END);
		sb.append(TABLE_END);
		return sb.toString();
	}

	/**
	 * 画像位置指定ブロック開始。
	 */
	private static final String IMAGE_BLOCK_BEGIN =
			"				<fo:block-container position=\"absolute\" top=\"${top}pt\" left=\"${left}pt\" width=\"${width}pt\" height=\"${height}pt\">\n";

	/**
	 * 画像位置指定ブロック終了。
	 */
	private static final String IMAGE_BLOCK_END =
			"				</fo:block-container>\n";

	/**
	 * 画像の配置タグを作成します。
	 * @param tinfo テーブル情報。
	 * @return 画像の配置タグ。
	 */
	private String getImageXml(final TableInfo tinfo) {
		StringBuilder sb = new StringBuilder();
		for (ImageInfo iinfo: tinfo.getImageList()) {
			String imageBlockBegin = IMAGE_BLOCK_BEGIN;
			imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{top\\}", "" + iinfo.getTop());
			imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{left\\}", "" + iinfo.getLeft());
			imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{height\\}", "" + iinfo.getHeight());
			imageBlockBegin = imageBlockBegin.replaceAll("\\$\\{width\\}", "" + iinfo.getWidth());
			sb.append(imageBlockBegin);
			sb.append("					<fo:block><fo:external-graphic src=\"" + iinfo .getImageSrc() + "\" width=\"" + iinfo.getWidth() + "pt\" height=\"" + iinfo.getHeight() + "pt\" content-width=\"" + iinfo.getWidth() + "pt\" content-height=\"" + iinfo.getHeight() + "pt\" border-style=\"dotted\" border-width=\"thin\"/></fo:block>\n");
			sb.append(IMAGE_BLOCK_END);
		}
		return sb.toString();
	}


	/**
	 * A4縦用のページマスタ。
	 */
	private static final String PAGE_MASTER =
			"	<fo:layout-master-set>\n" +
			"		<fo:simple-page-master page-height=\"${width}\" page-width=\"${height}\" margin-top=\"0mm\" margin-left=\"0mm\" margin-right=\"0mm\" margin-bottom=\"0mm\" master-name=\"PageMaster\">\n" +
			"			<fo:region-body margin-top=\"${topMargin}pt\" margin-left=\"${leftMargin}pt\" margin-right=\"${rightMargin}pt\" margin-bottom=\"${bottomMargin}pt\"/>\n" +
			"		</fo:simple-page-master>\n" +
			"	</fo:layout-master-set>\n";


	/**
	 * ページマスタを取得します。
	 * @param height ページの高さ。
	 * @param width ページの幅。
	 * @param landscape 横置きフラグ。
	 * @return ページマスタタグ。
	 */
	private String getPageMaster(final String height, final String width, final boolean landscape) {
		String pageMaster = PAGE_MASTER;
		if (landscape) {
			pageMaster = pageMaster.replaceAll("\\$\\{width\\}", width);
			pageMaster = pageMaster.replaceAll("\\$\\{height\\}", height);
		} else {
			pageMaster = pageMaster.replaceAll("\\$\\{width\\}", height);
			pageMaster = pageMaster.replaceAll("\\$\\{height\\}", width);
		}
		return pageMaster;
	}

	/**
	 * ページマスタを取得します。
	 * @param wb ワークブック。
	 * @param sb ページマスタを追加する文字列バッファ。
	 */
	private void getPageMaster(final Workbook wb, final StringBuilder sb) {
		Sheet sh = wb.getSheetAt(getSheetIndex());
		String pageMaster = PAGE_MASTER;
		logger.debug("paperSize=" + sh.getPrintSetup().getPaperSize());
		if (sh.getPrintSetup().getPaperSize() == PrintSetup.A3_PAPERSIZE) {
			pageMaster = this.getPageMaster("420mm", "297mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.A4_PAPERSIZE) {
			pageMaster = this.getPageMaster("297mm", "210mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.A5_PAPERSIZE) {
			pageMaster = this.getPageMaster("210mm", "148mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.B4_PAPERSIZE) {
			pageMaster = this.getPageMaster("354mm", "250mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.B5_PAPERSIZE) {
			pageMaster = this.getPageMaster("257mm", "182mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.LETTER_PAPERSIZE) {
			pageMaster = this.getPageMaster("279.4mm", "215.9mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.TABLOID_PAPERSIZE) {
			pageMaster = this.getPageMaster("431.8mm", "279.4mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.LEGAL_PAPERSIZE) {
			pageMaster = this.getPageMaster("355.6mm", "215.9mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.STATEMENT_PAPERSIZE) {
			pageMaster = this.getPageMaster("215.9mm", "139.7mm", sh.getPrintSetup().getLandscape());
		} else if (sh.getPrintSetup().getPaperSize() == PrintSetup.EXECUTIVE_PAPERSIZE) {
			pageMaster = this.getPageMaster("266.7mm", "184.1mm", sh.getPrintSetup().getLandscape());
		} else {
			pageMaster = this.getPageMaster("297mm", "210mm", sh.getPrintSetup().getLandscape());
		}
		double topMargin = sh.getMargin(Sheet.TopMargin) * 72;
		double bottomMargin = sh.getMargin(Sheet.BottomMargin) * 72;
		double leftMargin = sh.getMargin(Sheet.LeftMargin) * 72;
		double rightMargin = sh.getMargin(Sheet.RightMargin) * 72;
		pageMaster = pageMaster.replaceAll("\\$\\{topMargin\\}", "" + topMargin);
		pageMaster = pageMaster.replaceAll("\\$\\{bottomMargin\\}", "" + bottomMargin);
		pageMaster = pageMaster.replaceAll("\\$\\{leftMargin\\}", "" + leftMargin);
		pageMaster = pageMaster.replaceAll("\\$\\{rightMargin\\}", "" + rightMargin);
		sb.append(pageMaster);
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
		this.getPageMaster(wb, sb);
		Font f = wb.getFontAt(0);
		String pageBegin = PAGE_BEGIN;
		pageBegin = pageBegin.replaceAll("\\$\\{fontName\\}", f.getFontName());
		pageBegin = pageBegin.replaceAll("\\$\\{fontPoint\\}", "" + f.getFontHeightInPoints());
		sb.append(pageBegin);
		sb.append(this.getTableXml(wb, tinfo));
		sb.append(this.getImageXml(tinfo));
		sb.append(this.vImageList.toString());
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
		TableInfo tinfo = this.getTableInfo(wb);
		String xml = this.getXSLFO(wb, tinfo);
		logger.debug("XLS-SO:\n" + xml);
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
		ExcelToXSLFO conv = new ExcelToXSLFO();
		try {
			conv.parseAargs(args);
			conv.convert();
		} catch (ArgException e) {
			// e.printStackTrace();
			System.out.println("excel2xslfo [options] excelfile fofile");
			System.out.println("options:");
			System.out.println("-s sheetidx");
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
}
