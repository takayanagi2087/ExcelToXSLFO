package exeltoxslfo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;

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
	private static Logger logger = Logger.getLogger(ExcelToXSLFO.class);
	
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
	 * 引数指定の例外。
	 *
	 */
	private class ArgException extends Exception {
		
	};

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
			this.getBorderStyleAttribute(attrib, "left", this.style.getBorderLeftEnum());
			if (this.bottomRightStyle == null) {
				this.getBorderStyleAttribute(attrib, "bottom", this.style.getBorderBottomEnum());
				this.getBorderStyleAttribute(attrib, "right", this.style.getBorderRightEnum());
			} else {
				this.getBorderStyleAttribute(attrib, "bottom", this.bottomRightStyle.getBorderBottomEnum());
				this.getBorderStyleAttribute(attrib, "right", this.bottomRightStyle.getBorderRightEnum());
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
			if (c != null) {
				String hexcolor = c.getARGBHex();
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
				Font f = this.workbook.getFontAt((short) fidx);
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
			} else 	if (this.style.getAlignmentEnum() == HorizontalAlignment.CENTER) {
				attrib.append(" text-align=\"center\"");
			} else if (this.style.getAlignmentEnum() == HorizontalAlignment.RIGHT) {
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
		
		
		
		
		/**
		 * セルのスタイル情報を設定します。
		 * @return セルスタイル。
		 */
		/*
		public CellStyle getCellStyle() {
			return this.style;
		}*/
		
		
		
	}
	
	/**
	 * 画像情報。
	 *
	 */
	private class ImageInfo {
		/**
		 * 画像ファイルの上端の座標(pt)。
		 */
		private double top = 0;
		/**
		 * 画像ファイルの左端の座標(pt)。
		 */
		private double left = 0;
		
		/**
		 * 画像の高さ(pt)。
		 */
		private double height = 0;
		
		/**
		 * 画像の幅(pt)。
		 */

		private double width = 0;

		/**
		 * 画像データ。
		 */
		private XSSFPictureData imageData = null;
		
		/**
		 * コンストラクタ。
		 * @param top 画像の上端の位置(pt)。
		 * @param left 画像の左端の位置(pt)。
		 * @param height 画像の高さ(pt)。
		 * @param width 画像の幅(pt)。
		 * @param data 画像データ。
		 */
		public ImageInfo(final double top, final double left, final double height, final double width, final XSSFPictureData data) {
			this.top = top;
			this.left = left;
			this.height = height;
			this.width = width;
			this.imageData = data;
			logger.debug("ImageInfo:" + "," + top + "," + left + "," + height + "," + width);
		}
		
		
		
		/**
		 * 画像ファイルの上端の座標(pt)を取得します。
		 * @return 画像ファイルの上端の座標(pt)。
		 */
		public double getTop() {
			return top;
		}
		
		/**
		 * 画像ファイルの左端座標(pt)を取得します。
		 * @return 画像ファイルの左端座標(pt)。
		 */
		public double getLeft() {
			return left;
		}

		/**
		 * 画像の高さ(pt)を取得します。
		 * @return 画像の高さ(pt)。
		 */
		public double getHeight() {
			return height;
		}

		/**
		 * 画像の幅(pt)を取得します。
		 * @return 画像の幅(pt)。
		 */
		public double getWidth() {
			return width;
		}
		
		/**
		 * Base64形式の画像ソースを取得します。
		 * @return 画像ソース。
		 */
		public String getImageSrc() {
			String ret = "data:" + this.imageData.getMimeType() + ";base64, ";
			byte [] img = this.imageData.getData();
			String encoded = Base64.getEncoder().encodeToString(img);
			return ret + encoded;
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
			int cols = this.getColums(sh);
			this.cellInfo = new CellInfo[rows][cols];
			for (int r = 0; r < rows; r++) {
				for (int c = 0; c < cols; c++) {
					this.cellInfo[r][c] = new CellInfo(wb);
					Cell cell = this.getCell(sh, r, c);
					if (cell != null) {
						this.cellInfo[r][c].setStyle(cell.getCellStyle());
						this.cellInfo[r][c].setValue(ExcelToXSLFO.this.getCellValue(cell, fe));
						if (cell.getCellTypeEnum() == CellType.FORMULA) {
							CellValue cv = fe.evaluate(cell);
							this.cellInfo[r][c].setCellType(cv.getCellTypeEnum());
						} else {
							this.cellInfo[r][c].setCellType(cell.getCellTypeEnum());
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
		 * 付属ファイルの保存ディレクトリを取得します。
		 * @return 付属ファイルの保存ディレクトリ。
		 */
/*		private String getFilesPath() {
			String files = ExcelToXSLFO.this.getXslFoFile() + ".files";
			File filesdir = new File(files);
			if (!filesdir.exists()) {
				filesdir.mkdirs();
			}
			return files;
		}*/
		
		/**
		 * 画像ファイルを保存します。
		 * @param pic 画像情報。
		 * @return 画像ファイルの保存ファイル名。
		 * @throws Exception 例外。
		 */
/*		private String saveImage(final XSSFPicture pic) throws Exception {
			String filesdir = this.getFilesPath();
			XSSFPictureData data = pic.getPictureData();
			byte[] img = data.getData();
			String type = data.getMimeType();
			logger.debug("type=" + type);
			String filename = pic.getShapeName() + "." + type.replaceAll("image/", "");
			FileOutputStream os = new FileOutputStream(filesdir + File.separatorChar  + pic.getShapeName() + "." + type.replaceAll("image/", ""));
			try {
				os.write(img);
			} finally {
				os.close();
			}
			File dir = new File(filesdir);
			return dir.getName() + "/" + filename;
		}*/
		
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
			Font f = wb.getFontAt((short) 0);
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
			"								<fo:block>";

	
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
		if (cell.getCellTypeEnum() == CellType.BLANK) {
			value = "";
		} else if (cell.getCellTypeEnum() == CellType.STRING) {
			value = cell.getStringCellValue();
		} else if (cell.getCellTypeEnum() == CellType.FORMULA) {
			value = fmt.formatCellValue(cell, fe);
		} else {
			value = fmt.formatCellValue(cell);
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
//					String value = this.getCellValue(cell);
					String value = ci.getValue();
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
		Font f = wb.getFontAt((short) 0);
		String pageBegin = PAGE_BEGIN;
		pageBegin = pageBegin.replaceAll("\\$\\{fontName\\}", f.getFontName());
		pageBegin = pageBegin.replaceAll("\\$\\{fontPoint\\}", "" + f.getFontHeightInPoints());
		sb.append(pageBegin);
		sb.append(this.getImageXml(tinfo));
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
