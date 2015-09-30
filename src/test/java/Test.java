import java.awt.Color;
import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import com.incesoft.tools.excel.ExcelRowIterator;
import com.incesoft.tools.excel.ReaderSupport;
import com.incesoft.tools.excel.WriterSupport;
import com.incesoft.tools.excel.support.CellFormat;
import com.incesoft.tools.excel.xlsx.ExcelUtils;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;

public class Test {

	public static void main(String[] args) {

		// check is office2007 or 03 version
//		ExcelUtils.getExcelExtensionName(new File("/12.xlsx"));

//		ReaderSupport rxs = ReaderSupport.newInstance(ReaderSupport.TYPE_XLSX, new File("in.xlsx"));
//		rxs.open();
//		ExcelRowIterator it = rxs.rowIterator();
//		while (it.nextRow()) {
//			System.out.println(it.getCellValue(0));
//		}
//		rxs.close();

		SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(new File("in.xlsx"));
		for(Iterator it = getXlsxIterator(workbook.getSheet(0, false)); it.hasNext();){
			System.out.println(it.next());
		}
//		WriterSupport wxs = WriterSupport.newInstance(WriterSupport.TYPE_XLSX, new File("/out.xlsx"));
//		// WriterSupport wxs = WriterSupport.newInstance(WriterSupport.TYPE_XLS,
//		// new File("/out.xls"));
//		wxs.open();
//		wxs.increaseRow();
//		for (int i = 0; i < 5; i++) {
//			wxs.increaseRow();
//			wxs.writeRow(new String[] { "floydd" + i }, new CellFormat[] { new CellFormat(
//					(i % 2 == 0 ? Color.PINK.getRGB() : Color.GREEN.getRGB()), -1, 0) });
//		}
//		wxs.close();
	}


	private static Iterator<Map<String, String>> getXlsxIterator(com.incesoft.tools.excel.xlsx.Sheet sheet) {
		final com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader reader = sheet.newReader();
		return new Iterator<Map<String, String>>() {
			int count = 0;
			com.incesoft.tools.excel.xlsx.Cell[] row = reader.readRow();

			@Override
			public boolean hasNext() {
				return row != null;
			}

			@Override
			public Map<String, String> next() {

				Map<String, String> mapRow = new HashMap<String, String>();
				mapRow.put("current_row", ""+count);
				int colNum = 0;
				for (com.incesoft.tools.excel.xlsx.Cell cell : row) {
					if(cell==null){
						colNum++;
						continue;
					}
					String key = "column" + colNum;
					String strValue = cell.getValue();
					mapRow.put(key, strValue);
					colNum++;
				}
				count++;
				row = reader.readRow();
				return mapRow;
			}
		};
	}
}
