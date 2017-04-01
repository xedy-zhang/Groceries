/**
 * 
 */
package com.xedy.groceries.xls;

import java.io.File;
import java.io.IOException;
import java.util.List;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * xls文件处理工具类。使用两种处理技术：<i>jxl</i>、<i>poi</i>。
 * 
 * @author xedy_zhang@126.com
 *
 */
public class XLSHelper {
	
	/**
	 * 将数据写入xls文件。如果<b><i>datas</i></b>超过<b>65536行256列</b>，则仅写入<b>65536行256列</b>条数据。超过<b>65536行256列</b>请使用{@link #writeXlsx}
	 * @param datas 待写入xls数据集。
	 * @param filePath 文件路径。文件不存在时将创建。
	 * @throws IOException 文件创建，写入异常
	 */
	public static void writeXls(List<List<String>> datas,String filePath) throws IOException{
		if(datas == null || datas.size() == 0){
			return;//无数据，则返回
		}
		File file = fileCreator(filePath);
		try {
			writeXlsWithJxl(file,datas);
		} catch (BiffException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 根据指定文件完整路径创建文件
	 * @param filePath 文件完整路径
	 * @return 已创建的文件
	 * @throws IOException 文件创建异常
	 */
	private static File fileCreator(String filePath) throws IOException{
		File file  = new File(filePath);
		if(!(file.exists() && file.isFile())){
			file.createNewFile();
		}
		return file;
	}
	
	/**
	 * 将数据写入xls文件。获取过程嵌入到xls创建过程。
	 * @param filePath 文件路径。文件不存在时将创建。
	 * @param processor 单元格处理器
	 * @throws IOException 文件创建异常
	 */
	public static void writeXls(String filePath,AddSheetCellProcessor processor) throws IOException{
		File file = fileCreator(filePath);
		try {
			writeXlsWithJxl(file,processor);
		} catch (BiffException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 将数据获取嵌入到xls创建过程
	 * @param file 已创建xls文件
	 * @param processor 外部扩展处理实现
	 * @throws IOException 数据写入异常
	 * @throws BiffException 
	 */
	private static void writeXlsWithJxl(File file , AddSheetCellProcessor processor) throws IOException, BiffException{
		WritableWorkbook wwb = null;
		WritableSheet sheet = null;
		if(file.length() == 0){
			wwb = Workbook.createWorkbook(file);
			sheet = wwb.createSheet("SHEET 1", 0);
			sheet.getSettings().setVerticalFreeze(1);
		}else{
			wwb = Workbook.createWorkbook(file,Workbook.getWorkbook(file));
			sheet = wwb.getSheet(0);
		}
		try {
			processor.doProcessor(sheet);
			wwb.write();
			wwb.close();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 使用jxl工具创建xls文件
	 * @param file xls文件对象
	 * @param datas 待写入xls文件数据集合
	 * @throws IOException 数据写入异常
	 * @throws BiffException 
	 */
	private static void writeXlsWithJxl(File file,final List<List<String>> datas) throws IOException, BiffException{
		writeXlsWithJxl(file, new AddSheetCellProcessor() {
			public void doProcessor(WritableSheet sheet) throws RowsExceededException, WriteException {
				int rows = sheet.getRows(),cols = sheet.getColumns();
				for (int i = 0; ( (i+rows) < 65536 && i < datas.size() ) ; i++) {
					List<String> data = datas.get(i);
					for (int x = 0 ; ( (cols+x) < 256 && x < data.size() ) ; x++) {
						//已有行数+新增数据行 : rows + i
						sheet.addCell(new Label(x, rows+i, data.get(x)));
					}
				}
			}
		});
	}
	
	/**
	 * 将数据写入xlsx文件。如果<b><i>datas</i></b>超过<b>1048576行16384列</b>，则仅写入<b>1048576行16384列</b>条数据。
	 * @param datas 待写入xls数据集。
	 * @param filePath 文件路径。文件不存在时将创建。
	 */
	public static void writeXlsx(List<List<String>> datas,String filePath){
		//TODO: 待实现POI技术创建XLSX文件
	}
}
/**
 * 增加sheet单元格
 * @author xedy_zhang@126.com
 *
 */
interface AddSheetCellProcessor{
	public void doProcessor(WritableSheet sheet) throws RowsExceededException, WriteException;
}
