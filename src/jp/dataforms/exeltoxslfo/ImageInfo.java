package jp.dataforms.exeltoxslfo;

import java.util.Base64;

import org.apache.poi.xssf.usermodel.XSSFPictureData;

/**
 * 画像情報。
 *
 */
public class ImageInfo {
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
