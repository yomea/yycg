package poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 官方例子读取大数据量xls文件没有内存溢出问题
 * Hssf的事件驱动
 * @author mrt
 *
 */
//需要实现HSSFListener接口
public class ReadExcelHSSFEvent implements HSSFListener {
	
	private SSTRecord sstrec;

	/**
	 * This method listens for incoming records and handles them as required.
	 * 
	 * @param record
	 *            The record that was found while reading.
	 */
	public void processRecord(Record record) {

		switch (record.getSid()) {
		//如果Excel的内容是数字，则输出一下的字符串
		// the BOFRecord can represent either the beginning of a sheet or the
		// workbook
//		Encountered workbook
//		New sheet named: sheet0
//		Encountered sheet reference
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Row found, first column at 0 last column at 10
//		Cell found with value 0.0 at row 0 and column 0
//		Cell found with value 1.0 at row 0 and column 1
//		Cell found with value 2.0 at row 0 and column 2
//		Cell found with value 3.0 at row 0 and column 3
//		Cell found with value 4.0 at row 0 and column 4
//		Cell found with value 5.0 at row 0 and column 5
//		Cell found with value 6.0 at row 0 and column 6
//		Cell found with value 7.0 at row 0 and column 7
//		Cell found with value 8.0 at row 0 and column 8
//		Cell found with value 9.0 at row 0 and column 9
//		Cell found with value 0.0 at row 1 and column 0
//		Cell found with value 1.0 at row 1 and column 1
//		Cell found with value 2.0 at row 1 and column 2
//		Cell found with value 3.0 at row 1 and column 3
//		Cell found with value 4.0 at row 1 and column 4
//		Cell found with value 5.0 at row 1 and column 5
//		Cell found with value 6.0 at row 1 and column 6
//		Cell found with value 7.0 at row 1 and column 7
//		Cell found with value 8.0 at row 1 and column 8
//		Cell found with value 9.0 at row 1 and column 9
//		Cell found with value 0.0 at row 2 and column 0
//		Cell found with value 1.0 at row 2 and column 1
//		Cell found with value 2.0 at row 2 and column 2
//		Cell found with value 3.0 at row 2 and column 3
//		Cell found with value 4.0 at row 2 and column 4
//		Cell found with value 5.0 at row 2 and column 5
//		Cell found with value 6.0 at row 2 and column 6
//		Cell found with value 7.0 at row 2 and column 7
//		Cell found with value 8.0 at row 2 and column 8
//		Cell found with value 9.0 at row 2 and column 9
//		Cell found with value 0.0 at row 3 and column 0
//		Cell found with value 1.0 at row 3 and column 1
//		Cell found with value 2.0 at row 3 and column 2
//		Cell found with value 3.0 at row 3 and column 3
//		Cell found with value 4.0 at row 3 and column 4
//		Cell found with value 5.0 at row 3 and column 5
//		Cell found with value 6.0 at row 3 and column 6
//		Cell found with value 7.0 at row 3 and column 7
//		Cell found with value 8.0 at row 3 and column 8
//		Cell found with value 9.0 at row 3 and column 9
//		Cell found with value 0.0 at row 4 and column 0
//		Cell found with value 1.0 at row 4 and column 1
//		Cell found with value 2.0 at row 4 and column 2
//		Cell found with value 3.0 at row 4 and column 3
//		Cell found with value 4.0 at row 4 and column 4
//		Cell found with value 5.0 at row 4 and column 5
//		Cell found with value 6.0 at row 4 and column 6
//		Cell found with value 7.0 at row 4 and column 7
//		Cell found with value 8.0 at row 4 and column 8
//		Cell found with value 9.0 at row 4 and column 9
//		Cell found with value 0.0 at row 5 and column 0
//		Cell found with value 1.0 at row 5 and column 1
//		Cell found with value 2.0 at row 5 and column 2
//		Cell found with value 3.0 at row 5 and column 3
//		Cell found with value 4.0 at row 5 and column 4
//		Cell found with value 5.0 at row 5 and column 5
//		Cell found with value 6.0 at row 5 and column 6
//		Cell found with value 7.0 at row 5 and column 7
//		Cell found with value 8.0 at row 5 and column 8
//		Cell found with value 9.0 at row 5 and column 9
//		Cell found with value 0.0 at row 6 and column 0
//		Cell found with value 1.0 at row 6 and column 1
//		Cell found with value 2.0 at row 6 and column 2
//		Cell found with value 3.0 at row 6 and column 3
//		Cell found with value 4.0 at row 6 and column 4
//		Cell found with value 5.0 at row 6 and column 5
//		Cell found with value 6.0 at row 6 and column 6
//		Cell found with value 7.0 at row 6 and column 7
//		Cell found with value 8.0 at row 6 and column 8
//		Cell found with value 9.0 at row 6 and column 9
//		Cell found with value 0.0 at row 7 and column 0
//		Cell found with value 1.0 at row 7 and column 1
//		Cell found with value 2.0 at row 7 and column 2
//		Cell found with value 3.0 at row 7 and column 3
//		Cell found with value 4.0 at row 7 and column 4
//		Cell found with value 5.0 at row 7 and column 5
//		Cell found with value 6.0 at row 7 and column 6
//		Cell found with value 7.0 at row 7 and column 7
//		Cell found with value 8.0 at row 7 and column 8
//		Cell found with value 9.0 at row 7 and column 9
//		Cell found with value 0.0 at row 8 and column 0
//		Cell found with value 1.0 at row 8 and column 1
//		Cell found with value 2.0 at row 8 and column 2
//		Cell found with value 3.0 at row 8 and column 3
//		Cell found with value 4.0 at row 8 and column 4
//		Cell found with value 5.0 at row 8 and column 5
//		Cell found with value 6.0 at row 8 and column 6
//		Cell found with value 7.0 at row 8 and column 7
//		Cell found with value 8.0 at row 8 and column 8
//		Cell found with value 9.0 at row 8 and column 9
//		Cell found with value 0.0 at row 9 and column 0
//		Cell found with value 1.0 at row 9 and column 1
//		Cell found with value 2.0 at row 9 and column 2
//		Cell found with value 3.0 at row 9 and column 3
//		Cell found with value 4.0 at row 9 and column 4
//		Cell found with value 5.0 at row 9 and column 5
//		Cell found with value 6.0 at row 9 and column 6
//		Cell found with value 7.0 at row 9 and column 7
//		Cell found with value 8.0 at row 9 and column 8
//		Cell found with value 9.0 at row 9 and column 9
//		done.
		case BOFRecord.sid:
			BOFRecord bof = (BOFRecord) record;
			if (bof.getType() == bof.TYPE_WORKBOOK) {
				System.out.println("Encountered workbook");
				// assigned to the class level member
			} else if (bof.getType() == bof.TYPE_WORKSHEET) {
				System.out.println("Encountered sheet reference");
			}
			break;
		case BoundSheetRecord.sid:
			BoundSheetRecord bsr = (BoundSheetRecord) record;
			System.out.println("New sheet named: " + bsr.getSheetname());
			break;
		case RowRecord.sid:
			RowRecord rowrec = (RowRecord) record;
			System.out.println("Row found, first column at "
					+ rowrec.getFirstCol() + " last column at "
					+ rowrec.getLastCol());
			break;
			//处理数字值的Excel
		case NumberRecord.sid:
			NumberRecord numrec = (NumberRecord) record;
			System.out.println("Cell found with value " + numrec.getValue()
					+ " at row " + numrec.getRow() + " and column "
					+ numrec.getColumn());
			break;
		// SSTRecords store a array of unique strings used in Excel.
		//处理字符串值的Excel
		case SSTRecord.sid:
			sstrec = (SSTRecord) record;
			for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
				System.out.println("String table value " + k + " = "
						+ sstrec.getString(k));
			}
			break;
		//处理字符串值的Excel
		case LabelSSTRecord.sid:
			LabelSSTRecord lrec = (LabelSSTRecord) record;
			System.out.println(lrec.getRow()+"String cell found with value "
					+ sstrec.getString(lrec.getSSTIndex()));
			break;
		}
		
		
		System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%");
		
	}

	/**
	 * Read an excel file and spit out what we find.
	 * 
	 * @param args
	 *            Expect one argument that is the file to read.
	 * @throws IOException
	 *             When there is an error processing the file.
	 */
	public static void main(String[] args) throws IOException {
		// create a new file input stream with the input file specified
		// at the command line
		FileInputStream fin = new FileInputStream("d:/workbook.xls");
		// create a new org.apache.poi.poifs.filesystem.Filesystem
		POIFSFileSystem poifs = new POIFSFileSystem(fin);
		// get the Workbook (excel part) stream in a InputStream
		InputStream din = poifs.createDocumentInputStream("Workbook");
		// construct out HSSFRequest object
		HSSFRequest req = new HSSFRequest();
		// lazy listen for ALL records with the listener shown above
		//添加一个事件驱动
		req.addListenerForAllRecords(new ReadExcelHSSFEvent());
		// create our event factory
		HSSFEventFactory factory = new HSSFEventFactory();
		// process our events based on the document input stream
		//开始缺乏事件驱动，调用public void processRecord(Record record)方法
		factory.processEvents(req, din);
		// once all the events are processed close our file input stream
		fin.close();
		// and our document input stream (don't want to leak these!)
		din.close();
		System.out.println("done.");
	}
}
