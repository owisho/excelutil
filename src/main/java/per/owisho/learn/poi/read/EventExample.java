package per.owisho.learn.poi.read;

import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;

public class EventExample implements HSSFListener{

	private SSTRecord sstrec;
	
	@Override
	public void processRecord(Record record) {
		switch(record.getSid()) {
		
			//the BOFRecord can represent either the beginning of a sheet or the workbook
		case BOFRecord.sid:
			BOFRecord bof = (BOFRecord)record;
			if(bof.getType()==BOFRecord.TYPE_WORKSHEET) {
				System.out.println("Encountered workbook");
				//assigned to the class level member
			}else if(bof.getType()==BOFRecord.TYPE_WORKSHEET) {
				System.out.println("Encountered sheet reference");
			}
			break;
		case BoundSheetRecord.sid:
			BoundSheetRecord bsr = (BoundSheetRecord)record;
			System.out.println("New sheet named:"+bsr.getSheetname());
			break;
		case RowRecord.sid:
			RowRecord rowrec = (RowRecord) record;
			System.out.println("Row found, first column at "+rowrec.getFirstCol()+" last column at "+rowrec.getLastCol());
			break;
		case NumberRecord.sid:
			NumberRecord numrec = (NumberRecord)record;
			System.out.println("Cell fount with value "+ numrec.getValue()+" at row "+numrec.getRow()+" and column " + numrec.getColumn());
			break;
		case SSTRecord.sid:
			sstrec = (SSTRecord)record;
			for(int k=0;k<sstrec.getNumUniqueStrings();k++) {
				System.out.println("String table value "+k+" = "+sstrec.getString(k) );
			}
			break;
		case LabelSSTRecord.sid:
			LabelSSTRecord lrec = (LabelSSTRecord) record;
			System.out.println("String cell found with value "+sstrec.getString(lrec.getSSTIndex()));
			break;
		}
	}

}
