import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

enum  Airport_Format{
	ICAO,IATA;
}


public class TableTool {
	static boolean NullCell(Cell cell)
	{
		
		if(cell==null || cell.getCellTypeEnum()==CellType.BLANK )
			return true;
		return false;
	}
	static String getStringValue(Cell cell)
	{
		if(cell.getCellTypeEnum()==CellType.NUMERIC)
		{
			 if (DateUtil.isCellDateFormatted(cell)) {
                return ReadExcel.basicTimeFormat.format(cell.getDateCellValue()).toString();
             } else {
            	 Double original=(cell.getNumericCellValue());
            	 if(original-original.intValue() <0.0001)
            		 return String.valueOf(original.intValue());
            	 else
            		 return String.valueOf(original);
             }
		}
		
		else if(cell.getCellTypeEnum()==CellType.STRING)
			return cell.getRichStringCellValue().getString().trim();
		
		else if(cell.getCellTypeEnum()==CellType.BOOLEAN)
			return Boolean.toString(cell.getBooleanCellValue());
		
		else if(cell.getCellTypeEnum()==CellType.FORMULA)
			return cell.getCellFormula();
		
		//for blank,error
		return "";
	}
	
	@SuppressWarnings("deprecation")
	public static void setDateValue(Cell cell1,Cell cell2,ColRef ref,int offset ) throws InvocationTargetException, InterruptedException
	{
	
		int mid= ref.type.indexOf('_');
		String type=ref.type.substring(0,mid);
		String format=ref.type.substring(mid+1);
	
		if(type.equals("D"))
		{
			SimpleDateFormat dateformat = new SimpleDateFormat(format,Locale.US);
		
			Date date=cell1.getDateCellValue();
			date.setHours(date.getHours()+offset );
			
			cell2.setCellValue(dateformat.format(date).toString());
			
			//System.out.println("new date:"+dateformat.format(cell1.getDateCellValue()).toString());
		}else
			ReadExcel.printInfo(ReadExcel.InfoLevel.Error,"Wrong Refletion file:Date formate configured wrong:"+type);
		
	}
	@SuppressWarnings("deprecation")
	public static void setDateValue(String value, Cell cell2, SimpleDateFormat oldformat,String format,int offset) throws InvocationTargetException, InterruptedException {
		
		
		try {
			
			Date date=oldformat.parse(value);
			date.setHours(date.getHours()+offset );
			
			SimpleDateFormat newdateformat ;
			newdateformat = new SimpleDateFormat(format,Locale.US);	
			
			
			cell2.setCellValue(newdateformat.format(date).toString());
			
	} catch (ParseException e) {
		ReadExcel.printInfo(ReadExcel.InfoLevel.Error,"Error:string("+value+") with format"+oldformat.toString()+" to date parse error");
			e.printStackTrace();
		}
		
	}

}

