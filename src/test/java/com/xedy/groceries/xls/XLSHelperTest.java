package com.xedy.groceries.xls;

import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

import junit.framework.TestCase;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class XLSHelperTest extends TestCase {
	private String filePath="test.xls";
	private List<List<String>> datas = new ArrayList<List<String>>();
	
	@Override
	protected void setUp() throws Exception {
		super.setUp();
		for(int i = 0 ; i < 1000 ; i++){
			List<String> data = new ArrayList<String>();
			for(int x = 0 ; x < 10 ; x++){
				if(i==0){
					data.add("标题行"+x);
				}else{
					data.add(MessageFormat.format("row{0}-colum{1}", new String[]{String.valueOf(i),String.valueOf(x)}));
				}
			}
			datas.add(data);
		}
	}
	
	public void testWriteXls() {
		
		try {
			XLSHelper.writeXls(datas, filePath);
		} catch (IOException e) {
			e.printStackTrace();
		}
		assertEquals(true, true);
	}
	
	public void testWriteXls_DataEmbedded(){
		
		try {
			XLSHelper.writeXls("xls-embedded.xls",new AddSheetCellProcessor(){
				public void doProcessor(WritableSheet sheet) throws RowsExceededException, WriteException {
					sheet.addCell(new Label(0, sheet.getRows(), "嵌入xls创建过程..."));
				}
			});
		} catch (IOException e) {
			e.printStackTrace();
		}
		assertEquals(true, true);
	}
}
