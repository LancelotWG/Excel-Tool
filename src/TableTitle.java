import java.io.File;
import java.io.IOException;
import java.util.ArrayList;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

class Entry{
	//String excelname; //source excelfile

	int  colindex=0;     //source sheet index;
	int  backcolindex=0; //source sheet index;
	String  func=null;   //func
	String val=null;    //fix vaule;
	
	public boolean check()
	{
	   //at least has an validate value
		if(colindex==ResultRef.Blank && backcolindex==ResultRef.Blank && func==null && val==null)
			return false;
		
		return true;
	}
	public void print()
	{
		//System.out.print("  excelname:"+excelname  +" colindex"+ colindex);
		System.out.print(",colindex="+ colindex + ",back_colindex="+ backcolindex);
		
		if(func!=null)
			System.out.print(",func="+func);
				
		if(val!=null)
			System.out.print(",defaultVal="+val);		
	}	
	
}
class ColRef{
	int    resultcol=0;  //result file col index
	String type;      //reslult col data type,supprort date
	String val=null;  //default value
	
	int basic_index;  //source:basic index value
	
	boolean hasabnormal=false;
	boolean  hasspecial=false;
	Entry special;  //indicate special handle,and the excelname col is "N"
	Entry abnormal;

	public boolean check(int i)
	{
		if(resultcol==ResultRef.Blank)
		{
			System.out.println("Ref Row "+ (i+1)+ "failed to check!");
			return false;
		}
		
		return true;
	}
	public void println()
	{
		System.out.print("resultcol="+resultcol + ",basic_index="+basic_index);	
		if(type!=null)
			System.out.print(",Type="+type);
		else
			System.out.print(",Type=null");
		
		
		if(val!=null)
			System.out.print(",DefaultVal="+val);
		else
			System.out.print(",DefaultVal=null");
		
		if(hasabnormal)
			abnormal.print();
		
		
		if(hasspecial)
			special.print();
	
		
		System.out.println();
	}
	
}




class ResultRef {

	ArrayList<ColRef> basiclist =new ArrayList<ColRef>();
	ArrayList<ColRef> delaylist =new ArrayList<ColRef>();
	ArrayList<ColRef> cancellist =new ArrayList<ColRef>();
	ArrayList<ColRef> swaplist =new ArrayList<ColRef>();	
	static int Blank=-1;
	
	static int ref_sheet=0;
	static int val_sheet=1;
	
	static String ref_sheet_name="Ref";
	static String val_sheet_name="Value";
	
	static int ref_Sheet_Header_Row=1;
	static int ref_Sheet_Footer_Row=0;

	
	static int value_Sheet_Header_Row=1;
	static int value_Sheet_Footer_Row=0;
	
	private static Logger logger = Logger.getLogger(ResultRef.class);  
	
	public boolean init_resultRef(String refname ) throws IOException, InvalidFormatException 
	{
		
		boolean flag =true;	
		XSSFWorkbook wb= new XSSFWorkbook(OPCPackage.open(new File(refname)));			
		
		Sheet ref = wb.getSheetAt(ref_sheet);				
		if(!ref.getSheetName().equals(ref_sheet_name))
		{
			wb.close();	
			logger.error("Sheet("+ref_sheet+") of ref excel file is not named"+val_sheet_name);
			return false;
		}
		
		Sheet val=wb.getSheetAt(val_sheet);
		if(!val.getSheetName().equals(val_sheet_name))
		{
			logger.error("Sheet("+ref_sheet+") of ref excel file is not named"+val_sheet_name);			
			wb.close();
			return false;
		}
		
		
		//handle val table sheet
		handleValTableSheet(val);
		
		Cell cell_r_c, cell_r_t, cell_r_default;
		Cell cell_b_colindex;
		Cell cell_excelname,cell_colindex,cell_backcolindex,cell_func,cell_fix;
		Row row;
		String excelname;
		
	
		for (int j=ref_Sheet_Header_Row;j<=ref.getLastRowNum()-ref_Sheet_Footer_Row;j++)
	    {
		
			row=ref.getRow(j);
			ColRef entry=new ColRef();

			cell_r_c=row.getCell(0);
		    entry.resultcol = (int) cell_r_c.getNumericCellValue();	
		    
		    cell_r_t=row.getCell(1);
		    if(!TableTool.NullCell(cell_r_t) )
		    	entry.type=cell_r_t.getStringCellValue().trim();
		    else
		    	entry.type=null;
		    
		    
		    cell_r_default=row.getCell(2);
		    if(!TableTool.NullCell(cell_r_default ))
		    	entry.val=cell_r_default.getStringCellValue().trim();
		    else
		    	entry.val=null;
		    
		    
		    cell_b_colindex=row.getCell(3);
		    if(!TableTool.NullCell(cell_b_colindex ))
		    	entry.basic_index=(int) cell_b_colindex.getNumericCellValue();
		    else
		    	 entry.basic_index=Blank;	

		    cell_excelname=row.getCell(4);
			cell_colindex=row.getCell(5);
			cell_backcolindex=row.getCell(6);
			cell_func=row.getCell(7);
			cell_fix=row.getCell(8);
		    if(!TableTool.NullCell(cell_excelname) )
		    {
		    	
		    	excelname=cell_excelname.getRichStringCellValue().getString().trim(); //source excelfile
		    	if(excelname.equals("N"))
		    	{
		    		entry.hasspecial=true;
		    		entry.special=handleSpecial(cell_colindex,cell_backcolindex,cell_func,cell_fix);
		    	}else{
		    		entry.hasabnormal=true;
			    	entry.abnormal=handleAbnormal(cell_colindex,cell_backcolindex,cell_func,cell_fix);
			    	
				    if(excelname.equals("D"))
				    	delaylist.add(entry);				    	
				    else if(excelname.equals("C"))
				    	cancellist.add(entry);
				    else if(excelname.equals("S"))
				    	swaplist.add(entry);
				    else
				    {
				    	System.out.println(" error excelname in ref excel file");
				    	logger.error(excelname+" in row " +(j+1)+" in ref excel file  is not a correct file name configure");
				    }

		    	}
		    	
		    	basiclist.add(entry);
		    }
		    else if(entry.check(j))
		    	basiclist.add(entry);
		    
	    }

	    wb.close();
	    return flag;
	
	}
	
	
	private void  handleValTableSheet(Sheet sheet) 
	{
		
		
		Cell nameCell,valueCell,commentsCell;
		Row row;
		String name;
		
		for (int j=value_Sheet_Header_Row;j<=sheet.getLastRowNum()-value_Sheet_Footer_Row;j++)
		{
			row=sheet.getRow(j);
			nameCell=row.getCell(0);			
			valueCell=row.getCell(1);
			commentsCell=row.getCell(2);
			
			name=nameCell.getStringCellValue().trim();
			if(name==null ||name.equals(""))
				continue;
			name=name.toLowerCase();
			
			int intvalue=0;
			String stringvalue="";
			if(valueCell.getCellTypeEnum()==CellType.NUMERIC)
				intvalue=(int)valueCell.getNumericCellValue();
			else if(valueCell.getCellTypeEnum()==CellType.STRING)
				stringvalue=valueCell.getRichStringCellValue().getString().trim();
			
			switch(name){
				//airport code format ----------------------
				case "airport_name_format":
				{
					String value=stringvalue;
					if(value.equals("ICAO"))
					{
						ReadExcel.airport_Format=Airport_Format.ICAO;
						ReadExcel.airport_Format_dir=commentsCell.getStringCellValue().trim();					
					}else if(value.equals("IATA")){
						ReadExcel.airport_Format=Airport_Format.IATA;
						ReadExcel.airport_Format_dir=null;
					}	
					break;
				}
				case "airport_format1_col":
				{					
					ReadExcel.airport_format1_col=intvalue;
					break;
				}
				case "airport_format2_col":
				{
					ReadExcel.airport_format2_col=intvalue;
					break;
				}
				case "airportcode_header_row":
				{
					ReadExcel.airportCode_Header_Row=intvalue;
					break;
				}
				case "airportcode_footer_Row":
				{
					ReadExcel.airportCode_Footer_Row=intvalue;
					break;
				}//result sheet setting----------------------
				case "result_sheet_name":
				{
					ReadExcel.result_sheet_name=stringvalue;
					break;
				}
				case "result_sheet":
				{
					ReadExcel.result_sheet=intvalue;				
					break;
				}//basic sheet setting----------------------
				case "basic_sheet":
				{
					ReadExcel.basic_Sheet=intvalue;
					break;
				}
				case "basic_sheet_name":
				{
					ReadExcel.basic_Sheet_Name=stringvalue;
					break;
				}
				case "basic_header_row":
				{
					ReadExcel.basic_Header_Row=intvalue;
					break;
				}case "basic_footer_row":
				{
					ReadExcel.basic_Footer_Row=intvalue;
					break;
				}
				case "basic_flight_col":
				{
					ReadExcel.basic_Fight_Col=intvalue;
					break;
				}
				case "basic_departure_col":
				{
					ReadExcel.basic_Departure_Col=intvalue;
					break;
				}
				case "basic_takeoff_col":
				{
					ReadExcel.basic_Takeoff_col=intvalue;
					break;
				}
				case "basic_arrival_col":
				{
					ReadExcel.basic_Arrival_col=intvalue;
					break;
				}				
				//basic aircraft data sheet
				case "basic_f_t_sheet":
				{
					ReadExcel.basic_F_T_sheet=intvalue;
					break;
				}
				case "basic_f_t_sheet_name":
				{					
					ReadExcel.basic_F_T_sheet_name=stringvalue;
					break;
				}
				case "basicaircraft_tailnumbercol":
				{
					ReadExcel.basicAircraft_TailNumberCol=intvalue;
					break;
				}
				case "basic_aircraftdata_header_row":
				{
					ReadExcel.basic_AircraftData_Header_Row=intvalue;
					break;
				}case "basic_aircraftdata_footer_row":
				{
					ReadExcel.basic_AircraftData_Footer_Row=intvalue;
					break;
				}
				case "basicaircraft_fleetidcol":
				{
					ReadExcel.basicAircraft_FleetIDCol=intvalue;
					break;
				}//delay excel sheet data				
				case "delay_sheet":
				{
					ReadExcel.delay_sheet=intvalue;
					break;
				}
				case "delay_sheet_name":
				{
					ReadExcel.delay_sheet_name=stringvalue;
					break;
				}
				case "delay_hour_offset":
				{
					ReadExcel.delay_Hour_Offset=intvalue;
					break;
				}
				case "delay_minute_offset":
				{
					ReadExcel.delay_Minute_Offset=intvalue;
					break;
				}
				case "delay_fight_col":
				{
					ReadExcel.delay_Fight_Col=intvalue;
					break;
				}
				case "delay_departure_col":
				{
					ReadExcel.delay_Departure_Col=intvalue;
					break;
				}
				case "delay_takeoff_col":
				{
					ReadExcel.delay_Takeoff_col=intvalue;
					break;
				}
				case "delay_arrival_col":
				{
					ReadExcel.delay_Arrival_col=intvalue;
					break;
				}
				case "delay_header_row":
				{
					ReadExcel.delay_Header_Row=intvalue;
					break;
				}
				case "delay_footer_row":
				{
					ReadExcel.delay_Footer_Row=intvalue;
					break;
				}//cance excel sheet 
				case "cancel_sheet":
				{
					ReadExcel.cancel_sheet=intvalue;
					break;
				}
				case "cancel_sheet_name":
				{
					ReadExcel.cancel_sheet_name=stringvalue;
					break;
				}
				case "cancel_hour_offset":
				{
					ReadExcel.cancel_Hour_Offset=intvalue;
					break;
				}
				case "cancel_minute_offset":
				{
					ReadExcel.cancel_Minute_Offset=intvalue;
					break;
				}
				case "cancel_fight_col":
				{
					ReadExcel.cancel_Fight_Col=intvalue;
					break;
				}
				case "cancel_departure_col":
				{
					ReadExcel.cancel_Departure_Col=intvalue;
					break;
				}
				case "cancel_takeoff_col":
				{
					ReadExcel.cancel_Takeoff_col=intvalue;
					break;
				}
				case "cancel_arrival_col":
				{
					ReadExcel.cancel_Arrival_col=intvalue;
					break;
				}
				case "cancel_header_row":
				{
					ReadExcel.cancel_Header_Row=intvalue;
					break;
				}
				case "cancel_footer_row":
				{
					ReadExcel.cancel_Footer_Row=intvalue;
					break;
				}//swap sheet 1					
				case "swap_sheet1":
				{
					ReadExcel.swap_sheet1=intvalue;
					break;
				}
				
				case "swap_sheet_name1":
				{
					ReadExcel.swap_sheet_name1=stringvalue;
					break;
				}
				case "swap_fight_col1":
				{
					ReadExcel.swap_Fight_Col1=intvalue;
					break;
				}
				case "swap_departure_col1":
				{
					ReadExcel.swap_Departure_Col1=intvalue;
					break;
				}
				case "swap_departure_hour_col1":
				{					
					ReadExcel.swap_Departure_Hour_Col1=intvalue;
					break;
				}
				case "swap_takeoff_col1":
				{
					ReadExcel.swap_Takeoff_col1=intvalue;
					break;
				}
				
				case "swap_arrival_col1":
				{
					ReadExcel.swap_Arrival_col1=intvalue;
					break;
				}
				case "swap_aircraft_col1":
				{
					ReadExcel.swap_Aircraft_Col1=intvalue;
					break;
				}//swap sheet2 data
				case "swap_sheet2":
				{
					ReadExcel.swap_sheet2=intvalue;
					break;
				}
				
				case "swap_sheet_name2":
				{
					ReadExcel.swap_sheet_name2=stringvalue;
					break;
				}
				case "swap_fight_col2":
				{
					ReadExcel.swap_Fight_Col2=intvalue;
					break;
				}
				case "swap_departure_col2":
				{
					ReadExcel.swap_Departure_Col2=intvalue;
					break;
				}
				case "swap_departure_hour_col2":
				{
					ReadExcel.swap_Departure_Hour_Col2=intvalue;
					break;
				}
				case "swap_takeoff_col2":
				{
					ReadExcel.swap_Takeoff_col2=intvalue;
					break;
				}
				case "swap_arrival_col2":
				{
					ReadExcel.swap_Arrival_col2=intvalue;
					break;
				}
				case "swap_aircraft_col2":
				{
					ReadExcel.swap_Aircraft_Col2=intvalue;
					break;
				}//swap common data
				case "swap_header_row":
				{
					ReadExcel.swap_Header_Row=intvalue;
					break;
				}
				case "swap_footer_row":
				{
					ReadExcel.swap_Footer_Row=intvalue;
					break;
				}
				case "swap_hour_offset":
				{
					ReadExcel.swap_Hour_Offset=intvalue;
					break;
				}
				case "swap_minute_offset":
				{
					ReadExcel.swap_Minute_Offset=intvalue;
					break;
				}
				default:
				{
					System.out.println(" unknown configure name in ref value sheet:"+ name);
					logger.warn("unknown configure name in ref value sheet:"+ name+" in row "+(j+1));
					break;
				}				
			
			}

		}
			
	}
	



	public Entry handleSpecial(Cell index1,Cell index2,Cell func,Cell fix)
	{		
		return handleAbnormal(index1,index2,func,fix);	
	}
	
	
	public Entry handleAbnormal(Cell index1,Cell index2,Cell func,Cell fix)
	{

		Entry entry =new Entry();
	    if(TableTool.NullCell(index1))
	    	 entry.colindex=Blank;
	    else
	    	entry.colindex = (int) index1.getNumericCellValue();	
	    
	    if(TableTool.NullCell(index2))
	    	 entry.backcolindex=Blank;
	    else
	    	entry.backcolindex = (int) index2.getNumericCellValue();	
	    
	    
	    if(TableTool.NullCell(func))
	    	 entry.func=null;
	    else
	    	entry.func=func.getRichStringCellValue().getString().trim();

	    if(TableTool.NullCell(fix))
	    	 entry.val=null;
	    else
	    	entry.val=fix.getRichStringCellValue().getString().trim();
	    
		return entry;
	}
	
	
}




