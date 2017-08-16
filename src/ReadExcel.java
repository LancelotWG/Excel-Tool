
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Calendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import RecoveryPostSwap.SwapEntity;
import RecoveryPostSwap.SwapsPaser;
import javafx.scene.chart.PieChart.Data;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/*
		delay table :the first row, the last two row are cancel
		cancle table :the first row, the last row are cancel
		swap table:
*/
public class ReadExcel {

	static boolean Debug = false;
	static String String_Empty = "Null";
	static SimpleDateFormat swapTimeFormat = new SimpleDateFormat("ddMMYYYY");
	private static Logger logger = Logger.getLogger(ReadExcel.class);

	// basic
	static int basic_Sheet = 8;
	static String basic_Sheet_Name = "8_Flight_Data";

	static int basic_Hour_Offset = 0;// don't set in the ref configure sheet
	static int basic_Minute_Offset = 0;// don't set in the ref configure sheet

	static int basic_Fight_Col = 0;
	static int basic_Tail_Col = 6;
	static int basic_Departure_Col = 2; // format :14122016_0955
	static int basic_Arrival_Col = 3; // format :14122016_0955
	static int basic_Takeoff_col = 4;
	static int basic_Arrival_col = 5;

	static int basic_Header_Row = 1;
	static int basic_Footer_Row = 0;
	
	static int broken_Header_Row = 2;

	static SimpleDateFormat basicTimeFormat = new SimpleDateFormat("ddMMyyyy_HHmm");
	static String BasicTimeFormatString = "ddMMYYYY_HHmm";

	static int basic_F_T_sheet = 7;
	static String basic_F_T_sheet_name = "7_Aircraft_Data";
	static int basicAircraft_TailNumberCol = 0;
	static int basicAircraft_FleetIDCol = 1;
	static int basic_AircraftData_Header_Row = 1;
	static int basic_AircraftData_Footer_Row = 0;

	static HashMap<String, String> f_tail = new HashMap<String, String>();
	static HashMap<RowHash, Integer> f_row = new HashMap<RowHash, Integer>();
	static HashMap<String, Boolean> flightLocationMap = new HashMap<String, Boolean>();
	static HashSet<Flights> f_flights = new HashSet<Flights>();
	static ArrayList<Flights> f_joint_flights = new ArrayList<Flights>();
	static ArrayList<BrokenRecords> broken_records = new ArrayList<BrokenRecords>();

	// delay
	static int delay_sheet = 0;
	static String delay_sheet_name = "delayed-flights";

	static int delay_Hour_Offset = 8;
	static int delay_Minute_Offset = 15;
	static int delay_Pax_col = 12;
	static int delay_Fight_Col = 0;
	static int delay_Tail_Col = 1;
	static int delay_Departure_Col = 4;
	static int delay_Arrival_Col = 6;
	static int delay_Takeoff_col = 2;
	static int delay_Arrival_col = 3;
	static int delay_EDeparture_Col = 8;
	static int delay_EArrival_Col = 9;

	static int delay_Header_Row = 1;
	static int delay_Footer_Row = 2;

	// cancel
	static int cancel_sheet = 0;
	static String cancel_sheet_name = "cancelled-flights";

	static int cancel_Hour_Offset = 8;
	static int cancel_Minute_Offset = 15;

	static int cancel_Pax_col = 8;
	static int cancel_Fight_Col = 1;
	static int cancel_Tail_Col = 2;
	static int cancel_Departure_Col = 5;
	static int cancel_Arrival_Col = 7;
	static int cancel_Takeoff_col = 3;
	static int cancel_Arrival_col = 4;

	static int cancel_Header_Row = 1;
	static int cancel_Footer_Row = 1; ////// ?????????????????????

	// swap common
	static int swap_Hour_Offset = 8;
	static int swap_Minute_Offset = 15;
	static int swap_Header_Row = 2;
	static int swap_Footer_Row = 0;

	// turn around time limit
	static int turn_around_time_upperlimit = 20;
	static int turn_around_time_lowerlimit = 120;

	// swap1
	static int swap_sheet1 = 0;
	static String swap_sheet_name1 = "Within Subfleets";

	static int swap_Fight_Col1 = 0;
	static int swap_Tail_Col1 = 3;
	/*
	 * in swap file,the time is separate,Swap_Departure_col1 indicate the
	 * year_month_day, and the Swap_Departure_Hour_Col1 indicate the
	 * hour:hour_minute .
	 */
	static int swap_Departure_Col1 = 7;
	static int swap_Departure_Hour_Col1 = 8;
	static int swap_Takeoff_col1 = 5;
	static int swap_Arrival_col1 = 6;
	static int swap_Aircraft_Col1 = 4;
	static int swap_Pax_col1 = 10;

	// swap2
	static int swap_sheet2 = 1;
	static String swap_sheet_name2 = "Between Subleets";
	static int swap_Fight_Col2 = 0;
	static int swap_Departure_Col2 = 8;
	static int swap_Departure_Hour_Col2 = 9;
	static int swap_Takeoff_col2 = 6;
	static int swap_Arrival_col2 = 7;
	static int swap_Aircraft_Col2 = 5;

	// result

	static String result_sheet_name = "Final flight sheet";
	static int result_sheet = 0;

	static String broken_result_sheet_name = "Broken through flight ";
	static int broken_result_sheet = 0;

	// airport code table setting
	static HashMap<String, String> airport_Transform = new HashMap<String, String>();
	static Airport_Format airport_Format = Airport_Format.IATA;
	static String airport_Format_dir;
	static int airport_format1_col = 0;
	static int airport_format2_col = 1;
	static int airportCode_Header_Row = 1;
	static int airportCode_Footer_Row = 0;

	// letter (0-Z, 1-Y, 2-X, 3-W, 4-V, 5-U, 6-T, 7-S, 8-R, 9-Q)
	private static final HashMap<String, String> restore_flight_renaming_map;
	static {
		restore_flight_renaming_map = new HashMap<String, String>();
		restore_flight_renaming_map.put("0", "Z");
		restore_flight_renaming_map.put("1", "Y");
		restore_flight_renaming_map.put("2", "X");
		restore_flight_renaming_map.put("3", "W");
		restore_flight_renaming_map.put("4", "V");
		restore_flight_renaming_map.put("5", "U");
		restore_flight_renaming_map.put("6", "T");
		restore_flight_renaming_map.put("7", "S");
		restore_flight_renaming_map.put("8", "R");
		restore_flight_renaming_map.put("9", "Q");
	}

	private static void printValueConfigured() throws InvocationTargetException, InterruptedException {
		String info = "Below are the value configured in this run:" + "\n\tbasic_Header_Row=" + basic_Header_Row
				+ ",  basic_Footer_Row=" + basic_Footer_Row + "\n\tbasic_AircraftData_Header_Row="
				+ basic_AircraftData_Header_Row + ",  basic_AircraftData_Footer_Row=" + basic_AircraftData_Footer_Row

				+ "\n\tdelay_Header_Row=" + delay_Header_Row + ",delay_Footer_Row=" + delay_Footer_Row
				+ ",   delay_Hour_Offset=" + delay_Hour_Offset + ",  delay_Minute_Offset=" + delay_Minute_Offset
				+ "\n\tcancel_Header_Row=" + cancel_Header_Row + ",cancel_Footer_Row=" + cancel_Footer_Row
				+ ",  cancel_Hour_Offset=" + cancel_Hour_Offset + ",  cancel_Minute_Offset=" + cancel_Minute_Offset
				+ "\n\tswap_Header_Row=" + swap_Header_Row + ",  swap_Footer_Row=" + swap_Footer_Row
				+ ",  swap_Hour_Offset=" + swap_Hour_Offset + ",  swap_Minute_Offset=" + swap_Minute_Offset + "\n";

		printInfo(InfoLevel.Info, info);
	}

	// static public void handle( String basicname,String delayname,String
	// cancelname,String swapname,String resultname,String refname) throws
	// InvocationTargetException, InterruptedException
	static public void handle(String basicname, String abnormal, String resultname, String refname, String postSwapName, String brokenResult)
			throws InvocationTargetException, InterruptedException {
		try {
			readExcel(basicname, abnormal, resultname, refname, postSwapName, brokenResult);// delayname,
																				// cancelname,
																				// swapname
			// done
			if (Main.gui)
				JOptionPane.showMessageDialog(Main.frame, "done", "done", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {

			printInfo(InfoLevel.Error, e.getMessage());

			StackTraceElement stacktrace[] = e.getStackTrace();
			for (int i = 0; i < stacktrace.length; i++) {
				printInfo(InfoLevel.Error, "\t\t" + stacktrace[i].getClassName() + ":" + stacktrace[i].getMethodName()
						+ "(" + stacktrace[i].getLineNumber() + ")" + stacktrace[i].toString());
			}

			if (Main.gui == true) {
				JOptionPane.showMessageDialog(Main.frame, e.getMessage() + "\n For detail,see log file", "error",
						JOptionPane.INFORMATION_MESSAGE);
			} else {
				System.out.println("Serious Error:" + e.getMessage() + "\n For detail,see log file");
			}
		}

	}

	// String delayname,String cancelname,String swapname
	static private void readExcel(String basicname, String abnormal, String resultname, String refname,
			String postSwapName, String brokenResult)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {
		// clearHistory(); //If don't re-run,then,don't need to clear the
		// resource
		logger.info("\n\nNew Log.Below log items are create at: " + (new Date()));
		printInfo(InfoLevel.Info, "\nstarting cal...");

		/*
		 * ref handle,if there has error,then close the resource and return;
		 * 
		 */
		printInfo(InfoLevel.Info, "Checked the result reflection table:" + refname);
		File reffile = new File(refname);
		if (!reffile.exists()) {
			printInfo(InfoLevel.Error, "File(" + refname + ")is not exist.Program exit!");
			return;
		}
		ResultRef resultref = new ResultRef();
		if (!resultref.init_resultRef(refname)) {
			printInfo(InfoLevel.Error, "Reflection table " + refname
					+ "relation init error ,check the file format or make sure the file is exist!\nProgram exits!");
			return;
		}

		if (Debug) {
			for (int i = 0; i < resultref.basiclist.size(); i++)
				resultref.basiclist.get(i).println();
		}

		// print the value configure in value sheet of ref excel
		printInfo(InfoLevel.Info,
				"Reflection table:read  " + resultref.basiclist.size() + " items from ref excel table");
		printValueConfigured();

		/*
		 * check whether need do airport code format
		 * 
		 */
		if (airport_Format == Airport_Format.ICAO) {

			printInfo(InfoLevel.Info, "Init the Airport code format map data");
			if (airport_Format_dir == null || airport_Format_dir.equals("")) {
				printInfo(InfoLevel.Error,
						"Airport code format file:Airport format dir is empty,now we will use the IATA format instead!");
				airport_Format = Airport_Format.IATA;
			} else {
				airport_Format_dir = airport_Format_dir.trim();
				File file = new File(airport_Format_dir);
				if (!file.exists()) {
					printInfo(InfoLevel.Error,
							"Airport code format file:Airport format dir is wrong,now we will use the IATA format instead!");
					airport_Format = Airport_Format.IATA;
				} else {
					init_Airport_Format_data(file);
				}
			}
		}

		/*
		 * basic handle,if there has error,then close the resource and return;
		 * 
		 */
		printInfo(InfoLevel.Info, "Checked the Basic table data:" + basicname);
		File b = new File(basicname);
		if (!b.exists()) {
			printInfo(InfoLevel.Error, "basic table:" + basicname + " file is not exist!");
			return;
		}

		XSSFWorkbook resultwb = new XSSFWorkbook();
		Sheet result = resultwb.createSheet(result_sheet_name);

		handleBasic(basicname, result, resultref);

		if (postSwapName != null) {
			File postSwapFile = new File(postSwapName);
			if (postSwapName == null || !postSwapFile.exists()) {
				printInfo(InfoLevel.Info, "no post swap file does not exist.");
			} else {
				handlePostSwaps(postSwapName, basicname, f_row, result);
			}
		} else {
			printInfo(InfoLevel.Info, "no post swap file inputed");
		}

		/*
		 * abnormal excel file handle,if there has error,then close the resource
		 * and return;
		 * 
		 */
		File d = new File(abnormal);
		if (abnormal.equals("") || !d.exists())
			printInfo(InfoLevel.Error, "Abnormal file(" + abnormal + ")file is not exist!");
		else {
			printInfo(InfoLevel.Info, "Handel abnormal excel file  .....");
			handleAbnormal(abnormal, result, f_row, resultref, brokenResult);
		}

		// deal with post recovery window swap

		// flush result
		flushToFile(resultwb, result, resultname, brokenResult);
		resultwb.close();

	}

	static private boolean handlePostSwaps(String postSwapInput, String basicSheet, HashMap<RowHash, Integer> f_row,
			Sheet result) throws InvocationTargetException, InvalidFormatException, IOException, InterruptedException {
		// find all post swap aircraft
		SwapsPaser paser = new SwapsPaser();
		paser.paserData(postSwapInput);
		if (paser.swapList.size() < 1) {
			return false;
		}

		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date recoveryEnd787 = null;
		Date recoveryEnd737 = null;
		try {
			recoveryEnd737 = sdf.parse("2017-03-25 14:00:00");
			recoveryEnd787 = sdf.parse("2017-03-26 15:30:00");
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Map<String, Date> fleetRecoveryWindow;
		fleetRecoveryWindow = new HashMap<String, Date>();
		fleetRecoveryWindow.put("73", recoveryEnd737);
		fleetRecoveryWindow.put("78", recoveryEnd787);

		// Map<String, Date> aircraftAfterRecovery;
		java.util.Iterator<SwapEntity> it = paser.swapList.iterator();
		// System.out.println(paser.subfleetMap);
		while (it.hasNext()) {
			SwapEntity item = it.next();
			String subFleet = paser.subfleetMap.get(item.orgAircraft);
			if (null == subFleet || subFleet == "") {
				printInfo(InfoLevel.Info, "post swap- Ac has no fleet info " + item.orgAircraft);
				subFleet = "73";
				// continue;
			}
			item.recoveryEnd = fleetRecoveryWindow.get(subFleet);
			if (item.recoveryEnd == null) {
				printInfo(InfoLevel.Error, item.swappedAricraft + " aircraft has no subfleet " + item.orgAircraft);
				continue;
			}
			printInfo(InfoLevel.Info, item.swappedAricraft + " switches to the path of -> " + item.orgAircraft
					+ item.recoveryEnd.toString());

		}
		printInfo(InfoLevel.Info, "swaplist.count == " + paser.swapList.size());
		// find all flight should be swapped, swap
		return swapPostSwapFlihts(basicSheet, result, paser.swapList, f_row);
	}

	static private boolean swapPostSwapFlihts(String basicSheet, Sheet result, List<SwapEntity> swapList,
			HashMap<RowHash, Integer> rawHashTable)
			throws IOException, InvocationTargetException, InterruptedException, InvalidFormatException {
		printInfo(InfoLevel.Info, "Start  swap post recovery  swapPostSwapFlihts---");

		Cell cell, dptCell;
		Row row;
		Date dptTime = null;
		String tailNo;
		int postSwapCount = 0;
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd  HH:mm:ss");
		for (int i = basic_Header_Row; i <= result.getLastRowNum() - basic_Footer_Row; i++) {
			row = result.getRow(i);
			// if Departure time > recovery window time, and flight
			cell = row.getCell(23); // ac_reg
			if (null == cell) {
				printInfo(InfoLevel.Warn, "Post swap" + (i + 1) + " Col " + 24 + " is not date formate!");
				continue;
			}

			tailNo = cell.getStringCellValue().trim();
			tailNo = tailNo.replace("-", "");
			dptCell = row.getCell(9);
			if (dptCell.getCellTypeEnum() != CellType.STRING) {
				printInfo(InfoLevel.Warn, "Result Row " + (i + 1) + " Col " + 9 + dptCell.getCellType()
						+ " is not string! v= " + dptCell);
				continue;
			}

			try {
				dptTime = sdf.parse(dptCell.getStringCellValue());
			} catch (ParseException e) {
				// TODO Auto-generated catch block
				printInfo(InfoLevel.Warn, "Result Row " + (i + 1) + " Col " + 9 + dptCell.getCellType()
						+ " can't be converted to Date! v= " + dptCell);
				e.printStackTrace();
				continue;
			}
			// dptTime = row.getCell(9).getDateCellValue(); // STD

			java.util.Iterator<SwapEntity> it = swapList.iterator();
			while (it.hasNext()) {
				SwapEntity item = it.next();

				if (tailNo.equals(item.orgAircraft) && dptTime.after(item.recoveryEnd)) {
					postSwapCount++;
					printInfo(InfoLevel.Info, "post recovery swap -" + item.orgAircraft + " -" + item.swappedAricraft
							+ "-" + dptTime.toString());
					// maybe we should reuse refreshSwapRow fucn
					// value = value.substring(0, 1) + "-" + value.substring(1);
					cell.setCellValue(item.swappedAricraft.substring(0, 1) + "-" + item.swappedAricraft.substring(1));
				}
			}
		}
		printInfo(InfoLevel.Info, "Total Post Swap " + postSwapCount);
		return false;

	}

	/*
	 * private static void clearHistory() throws InvocationTargetException,
	 * InterruptedException {
	 * printInfo(InfoLevel.Info,"-------Clear history data-------------");
	 * airport_Transform.clear(); f_tail.clear(); f_row.clear();
	 * printInfo(InfoLevel.Info,"History file has been cleaned " ); }
	 */

	static private boolean handleBasic(String basicname, Sheet result, ResultRef resultref)
			throws IOException, InvocationTargetException, InterruptedException, InvalidFormatException {

		XSSFWorkbook basicwb = new XSSFWorkbook(OPCPackage.open(basicname));
		printInfo(InfoLevel.Info, "Init the aircraft map data" + basic_F_T_sheet_name);
		if (!getF_T_Map(basicwb, basic_F_T_sheet, basic_F_T_sheet_name, f_tail)) {
			printInfo(InfoLevel.Error, "Program exits,cause aircraft map data init error");
			basicwb.close();
			return false;
		}

		Sheet basic = basicwb.getSheetAt(basic_Sheet);
		if (!basic.getSheetName().equals(basic_Sheet_Name)) {
			printInfo(InfoLevel.Error,
					"basic table:" + basic_Sheet + "sheet of basic table is not named " + basic_Sheet_Name);
			basicwb.close();
			return false;
		}

		ColRef ref;
		Cell cell1, cell2;

		Row source, sink;
		for (int i = basic_Header_Row; i <= basic.getLastRowNum() - basic_Footer_Row; i++) {
			source = basic.getRow(i);
			sink = result.createRow(i);

			Cell timeCell = source.getCell(basic_Departure_Col);
			String day;
			if (null != timeCell) {
				day = source.getCell(basic_Departure_Col).getStringCellValue().trim();
				// String
				// day=TableTool.getStringValue(source.getCell(Basic_Departure_Col));
				day = day.substring(0, 8);// DDMMYYYY_HHMM,just need the
											// day,ignore
											// the hour and minut
			} else {
				printInfo(InfoLevel.Error, "Basic table: row" + (i + 1) + "Column" + basic_Departure_Col + "is null");
				continue;
			}

			String flight = String.valueOf((int) source.getCell(basic_Fight_Col).getNumericCellValue());

			String departure = source.getCell(basic_Takeoff_col).getStringCellValue().trim();
			String arrival = source.getCell(basic_Arrival_col).getStringCellValue().trim();
			// 修改自兰望桂
			/**
			 * 寻找联程航班
			 */
			String arrivalTime = source.getCell(basic_Arrival_Col).getStringCellValue().trim();
			String departureTime = source.getCell(basic_Departure_Col).getStringCellValue().trim();
			String tail = source.getCell(basic_Tail_Col).getStringCellValue().trim();
			Number number = new Number(flight);
			Place place = new Place(tail, arrival, departure, arrivalTime, departureTime);
			Flights flights = new Flights(number, place);
			f_flights.add(flights);
			// 修改自兰望桂

			RowHash rowentity = new RowHash(flight, day, departure, arrival);

			if (f_row.get(rowentity) != null)
				printInfo(InfoLevel.Error,
						"Basic table: row" + (i + 1) + " has two row with same key:" + rowentity.toString());
			else
				f_row.put(rowentity, new Integer(i));

			// record the row simple info in the basic:
			for (int j = 0; j < resultref.basiclist.size(); j++) {
				ref = resultref.basiclist.get(j);

				// don't write any thing
				if (ref.basic_index == ResultRef.Blank && ref.val == null && ref.hasspecial == false)
					continue;

				// special handle,and the excelname col is "N"
				if (ref.hasspecial) // indicate special handle,and the excelname
									// col is "N"
				{
					cell2 = sink.createCell(ref.resultcol);

					String hash = TableTool.getStringValue(source.getCell(ref.special.colindex));
					String specialVal = getSpecialVal(hash, ref.special.func);
					if (specialVal.equals(String_Empty)) {
						continue;
					}

					cell2.setCellValue(specialVal);
					continue;
				}

				// write default value
				if (ref.val != null)// set defalut value
				{
					cell2 = sink.createCell(ref.resultcol);
					cell2.setCellValue(ref.val);
					continue;
				}

				// write according to the ref
				cell1 = source.getCell(ref.basic_index);
				cell2 = sink.createCell(ref.resultcol);
				switch (cell1.getCellTypeEnum()) {
				case STRING:
					// for basic cell is a date value but not in date
					// format,meaning cell1.celltype==string
					if (ref.type != null && !ref.type.equals("")) {
						int mid = ref.type.indexOf('_');
						String type = ref.type.substring(0, mid);
						String format = ref.type.substring(mid + 1);
						if (type.equals("D")) {
							TableTool.setDateValue(cell1.getRichStringCellValue().getString().trim(), cell2,
									basicTimeFormat, format, basic_Hour_Offset);
						} else if (type.equals("S") && format.equalsIgnoreCase("-")) {
							String value = TableTool.getStringValue(cell1);
							value = value.substring(0, 1) + format + value.substring(1);
							cell2.setCellValue(value);
						}
					} else
						cell2.setCellValue(cell1.getRichStringCellValue().getString().trim());
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell1)) {
						// System.out.println(cell.getDateCellValue());
						TableTool.setDateValue(cell1, cell2, ref, basic_Hour_Offset);
					} else {
						// System.out.println(cell.getNumericCellValue());
						cell2.setCellValue(cell1.getNumericCellValue());
					}
					break;
				case BOOLEAN:
					// System.out.println(cell.getBooleanCellValue());
					cell2.setCellValue(cell1.getBooleanCellValue());
					break;
				case FORMULA:
					// System.out.println(cell1.getCellFormula());
					cell2.setCellValue(cell1.getCellFormula());
					break;
				default:
					// System.out.println();
					break;
				}

			}
		}
		// 修改自兰望桂
		/**
		 * 存储联程航班
		 */
		for (Iterator iterator = f_flights.iterator(); iterator.hasNext();) {
			Flights flights = (Flights) iterator.next();
			if (flights.places.size() >= 2) {
				f_joint_flights.add(flights);
			}
		}
		// 修改自兰望桂

		/*System.out.println("Size:" + f_flights.size());
		for (Iterator iterator = f_flights.iterator(); iterator.hasNext();) {
			Flights flights = (Flights) iterator.next();
			if (flights.places.size() >= 2) {
				System.out.println(flights.places.size());
				System.out.println("Flight:" + flights.number.flightNumber);
				ArrayList<Place> place = flights.places;
				for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
					Place place2 = (Place) iterator2.next();
					System.out.println("Tail:" + place2.tailNumber);
					System.out.println("Departure:" + place2.departure + " Time:" + place2.departureTime);
					System.out.println("Arrival:" + place2.arrival + " Time:" + place2.arrivalTime);
				}
			}
		}*/

		printInfo(InfoLevel.Info, "Handled Basic table:"
				+ (basic.getLastRowNum() + 1 - basic_Header_Row - basic_Footer_Row) + " row have been writed!");
		basicwb.close();

		return true;
	}

	static private void handleAbnormal(String abnormalname, Sheet result, HashMap<RowHash, Integer> f_row,
			ResultRef resultref, String brokenResult)
			throws InvalidFormatException, IOException, InvocationTargetException, InterruptedException {

		XSSFWorkbook abnormalwb = new XSSFWorkbook(OPCPackage.open(new File(abnormalname)));

		Sheet delay = abnormalwb.getSheetAt(delay_sheet);
		if (delay != null)
			handleDelay(delay, result, f_row, resultref);
		else
			printInfo(InfoLevel.Error, "Serous Error:sheet(" + delay_sheet + ") for deley data is not existsheet");

		Sheet cancel = abnormalwb.getSheetAt(cancel_sheet);
		if (cancel != null)
			handleCancel(cancel, result, f_row, resultref);
		else
			printInfo(InfoLevel.Error, "Serous Error:sheet(" + cancel_sheet + ") for cancel data is not exist");

		Sheet swap1 = abnormalwb.getSheetAt(swap_sheet1);
		if (swap1 != null)
			handleSwap1(swap1, result, f_row, resultref);
		else
			printInfo(InfoLevel.Error,
					"Serous Error:sheet(" + swap_sheet1 + ") for swap data(swaps within subfleets) is not exist");

		Sheet swap2 = abnormalwb.getSheetAt(swap_sheet2);
		if (swap2 != null)
			handleSwap2(swap2, result, f_row, resultref);
		else
			printInfo(InfoLevel.Error,
					"Serous Error:sheet(" + swap_sheet2 + ") for swap data(swaps between subfleets) is not exist");
		// change the name of restore flight
		handleBrokenFlight(abnormalname, resultref, brokenResult);
		// 修改自兰望桂
		String path = airport_Format_dir.trim();
		// 修改自兰望桂
		handleFlightLocation(path);
		boolean process_restore_flight = true;
		if (process_restore_flight) {
			handleRestoreFlight(delay, result, f_row, resultref);
		}

		abnormalwb.close();
	}

	/**
	 * 
	 * @param filePath
	 * @throws InvalidFormatException
	 * @throws IOException
	 * @author LancelotWG 2017/7/5
	 */
	static private void handleFlightLocation(String filePath) throws InvalidFormatException, IOException {
		XSSFWorkbook flightLocation = new XSSFWorkbook(OPCPackage.open(new File(filePath)));
		Sheet location = flightLocation.getSheetAt(1);
		int index = location.getLastRowNum();
		for (int i = delay_Header_Row; i <= location.getLastRowNum(); i++) {
			Row row = location.getRow(i);
			Cell IATACell = row.getCell(0);
			Cell countryCell = row.getCell(4);
			String IATA = IATACell.getStringCellValue();
			String country = "";
			if (countryCell.getCellTypeEnum() != CellType.STRING) {
				country = "null";
			} else {
				country = countryCell.getStringCellValue();
			}
			if (country.equals("China")) {
				flightLocationMap.put(IATA, true);
			} else {
				flightLocationMap.put(IATA, false);
			}
		}
		flightLocation.close();
	}

	@SuppressWarnings("deprecation")
	// not recognize international flights
	static private void handleRestoreFlight(Sheet delay, Sheet result, HashMap<RowHash, Integer> f_row,
			ResultRef resultref)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {
		if (!delay.getSheetName().equals(delay_sheet_name)) {
			printInfo(InfoLevel.Error,
					"Serous Error:delay table:" + (delay_sheet) + " sheet is not named " + delay_sheet_name);
			return;
		}
		Row sink;
		// System.out.println("testonly: " + delay_Footer_Row);
		for (int i = delay_Header_Row; i <= delay.getLastRowNum() - delay_Footer_Row; i++) {
			Row row = delay.getRow(i);

			// 机场起始站终点站
			Cell origLocationStationCell = row.getCell(2);
			String origLocationStation = origLocationStationCell.getStringCellValue();
			Cell destLocationStationCell = row.getCell(3);
			String destLocationStation = destLocationStationCell.getStringCellValue();
			// 机场起始站终点站

			Cell dateCell = row.getCell(delay_Departure_Col);
			if (null == dateCell || dateCell.getCellTypeEnum() != CellType.NUMERIC
					|| !DateUtil.isCellDateFormatted(dateCell)) {
				printInfo(InfoLevel.Warn,
						"Delay excel Row " + (i + 1) + " Col " + delay_Departure_Col + " is not date formate!");
				continue;
			}

			Date date = dateCell.getDateCellValue();
			date.setHours(date.getHours() + delay_Hour_Offset);
			date.setMinutes(date.getMinutes() + delay_Minute_Offset);
			String day = basicTimeFormat.format(date).toString().substring(0, 8); // just
																					// need
																					// day,ignore
																					// the
																					// hour
																					// and
																					// minute

			String flight = String.valueOf(((int) row.getCell(delay_Fight_Col).getNumericCellValue()));
			String departure = row.getCell(delay_Takeoff_col).getStringCellValue().trim();
			String arrival = row.getCell(delay_Arrival_col).getStringCellValue().trim();

			// rename the restore flight
			int dealy_est_col = 7;
			Cell etd_date_cell = row.getCell(dealy_est_col);
			if (etd_date_cell.getCellTypeEnum() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(etd_date_cell)) {
				printInfo(InfoLevel.Warn,
						"Delay excel Row " + (i + 1) + " Col " + dealy_est_col + " is not date formate!");
				continue;
			}
			Date etd_date = etd_date_cell.getDateCellValue();
			etd_date.setHours(etd_date.getHours() + delay_Hour_Offset);
			etd_date.setMinutes(etd_date.getMinutes() + delay_Minute_Offset);
			Calendar etd_calender = Calendar.getInstance();
			Calendar ori_calender = Calendar.getInstance();
			etd_calender.setTime(etd_date);
			ori_calender.setTime(date);
			int delayed_days = etd_calender.get(Calendar.DAY_OF_YEAR) - ori_calender.get(Calendar.DAY_OF_YEAR);
			// int delayed_days = etd_date.getDay() - date.getDate() ;
			// printInfo(InfoLevel.Warn,"tEST ONLY excel Row "+(i+1) +
			// etd_calender.get(Calendar.HOUR_OF_DAY) );
			// 5 hour is defined by Xiamen
			if ((delayed_days == 1 && etd_calender.get(Calendar.HOUR_OF_DAY) > 5) || (delayed_days > 1)) { // rename
																											// the
																											// flight
				printInfo(InfoLevel.Warn,
						"Delay excel Row " + (i + 1) + " Col " + dealy_est_col + " is a restore flight!");
				RowHash rowhash = new RowHash(flight, day, departure, arrival);
				Integer rowid = f_row.get(rowhash);
				if (rowid != null) {
					// find the row in result table;
					sink = result.getRow(rowid);

					Cell flight_output_cell = sink.getCell(2);
					if (flight_output_cell == null)
						continue;
					String new_flight_no = "";
					// 判断航班是国际航班还是国内航班
					Boolean o = flightLocationMap.get(origLocationStation);
					Boolean d = flightLocationMap.get(destLocationStation);
					if (flightLocationMap.get(origLocationStation) && flightLocationMap.get(destLocationStation)) {
						// 国内航班
						new_flight_no = gererateRestoreFlightNo(
								flight_output_cell.getRichStringCellValue().getString().trim(), delayed_days);
					} else {
						// 国际航班
						new_flight_no = gererateRestoreInternationalFlightNo(
								flight_output_cell.getRichStringCellValue().getString().trim(), delayed_days);
					}
					// 判断航班是国际航班还是国内航班
					flight_output_cell.setCellValue(new_flight_no);

				} else {
					printInfo(InfoLevel.Warn, "Delay Flight(row: " + (i + 1) + ")" + rowhash.toString()
							+ " doesn't have match row in basic table!");
				}
			}
		}
		printInfo(InfoLevel.Info, "Delay table:Handle "
				+ (delay.getLastRowNum() + 1 - delay_Header_Row - delay_Footer_Row) + " row in the Delay table");
		// delaywb.close();
	}

	// @SuppressWarnings("deprecation")
	static private String gererateRestoreFlightNo(String original_flight, int delayed_days) {
		String last_letter = original_flight.substring(original_flight.length() - 1);
		// last_letter = restore_flight_renaming_map.get(last_letter);
		String restore_flight = original_flight.substring(0, original_flight.length() - 1)
				+ restore_flight_renaming_map.get(last_letter);

		if (Debug)
			System.out.println(" Restore-fligth: O" + original_flight + " N:" + restore_flight);
		return restore_flight;
	}

	/**
	 * 
	 * @param original_flight
	 * @param delayed_days
	 * @return
	 * @author LancelotWG 2017/7/5
	 */
	static private String gererateRestoreInternationalFlightNo(String original_flight, int delayed_days) {
		String last_letter = original_flight.substring(original_flight.length() - 1);
		// last_letter = restore_flight_renaming_map.get(last_letter);
		String restore_flight = "";
		if (original_flight.length() < 6) {
			restore_flight = original_flight + "A";
		} else {
			restore_flight = original_flight.substring(0, original_flight.length() - 1)
					+ restore_flight_renaming_map.get(last_letter);
		}
		if (Debug)
			System.out.println(" Restore-fligth: O" + original_flight + " N:" + restore_flight);
		return restore_flight;
	}

	@SuppressWarnings("deprecation")
	static private void handleDelay(Sheet delay, Sheet result, HashMap<RowHash, Integer> f_row, ResultRef resultref)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {

		if (!delay.getSheetName().equals(delay_sheet_name)) {
			printInfo(InfoLevel.Error,
					"Serous Error:delay table:" + (delay_sheet) + " sheet is not named " + delay_sheet_name);
			return;
		}

		Row sink;
		/*
		 * System.out.println("delay row:"+delay.getLastRowNum());
		 * System.out.println("Delay_Footer_Row:"+Delay_Footer_Row);
		 */
		for (int i = delay_Header_Row; i <= delay.getLastRowNum() - delay_Footer_Row; i++) {
			Row row = delay.getRow(i);

			Cell dateCell = row.getCell(delay_Departure_Col);
			// printInfo(InfoLevel.Warn, "Delay excel Row "+(i+1) + " Col "+
			// delay_Departure_Col );
			if (dateCell == null || dateCell.getCellTypeEnum() != CellType.NUMERIC
					|| !DateUtil.isCellDateFormatted(dateCell)) {
				printInfo(InfoLevel.Warn,
						"Delay excel Row " + (i + 1) + " Col " + delay_Departure_Col + " is not date formate!");
				continue;
			}

			Date date = dateCell.getDateCellValue();
			date.setHours(date.getHours() + delay_Hour_Offset);
			date.setMinutes(date.getMinutes() + delay_Minute_Offset);
			String day = basicTimeFormat.format(date).toString().substring(0, 8); // just
																					// need
																					// day,ignore
																					// the
																					// hour
																					// and
																					// minute

			String flight = String.valueOf(((int) row.getCell(delay_Fight_Col).getNumericCellValue()));
			String departure = row.getCell(delay_Takeoff_col).getStringCellValue().trim();
			String arrival = row.getCell(delay_Arrival_col).getStringCellValue().trim();

			RowHash rowhash = new RowHash(flight, day, departure, arrival);

			Integer rowid = f_row.get(rowhash);

			if (rowid != null) {
				// find the row in result table;
				sink = result.getRow(rowid);
				refreshRow(row, sink, resultref.delaylist, delay_Hour_Offset);
			} else {
				printInfo(InfoLevel.Warn, "Delay Flight(row: " + (i + 1) + ")" + rowhash.toString()
						+ " doesn't have match row in basic table!");
			}

		}
		printInfo(InfoLevel.Info, "Delay table:Handle "
				+ (delay.getLastRowNum() + 1 - delay_Header_Row - delay_Footer_Row) + " row in the Delay table");
		// delaywb.close();

	}

	@SuppressWarnings("deprecation")
	static private void handleCancel(Sheet cancel, Sheet result, HashMap<RowHash, Integer> f_row, ResultRef resultref)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {

		/*
		 * XSSFWorkbook cancelwb=null; Sheet cancel=null; cancelwb=new
		 * XSSFWorkbook(OPCPackage.open(new File(cancelname)));
		 * cancel=cancelwb.getSheetAt(cancel_sheet);
		 * if(!cancel.getSheetName().equals(cancel_sheet_name)) {
		 * printInfo(InfoLevel.Error,"Serious Error: Cancel table"
		 * +cancel_sheet+" sheet of cancel table is not named "
		 * +cancel_sheet_name); cancelwb.close(); return; }
		 */
		if (!cancel.getSheetName().equals(cancel_sheet_name)) {
			printInfo(InfoLevel.Error, "Serious Error: Cancel table" + cancel_sheet
					+ " sheet of cancel table is not named " + cancel_sheet_name);
			// cancelwb.close();
			return;
		}
		Row sink;

		/*
		 * System.out.println("cancel row:"+cancel.getLastRowNum());
		 * System.out.println("cancel_Footer_Row:"+cancel_Footer_Row);
		 */
		for (int i = cancel_Header_Row; i <= cancel.getLastRowNum() - cancel_Footer_Row; i++) {
			Row row = cancel.getRow(i);

			Cell dateCell = row.getCell(cancel_Departure_Col);
			if (dateCell.getCellTypeEnum() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(dateCell)) {
				printInfo(InfoLevel.Warn,
						"Cancel excel:Row " + (i + 1) + " Col " + cancel_Departure_Col + " is not date formate!");
				continue;
			}

			Date date = dateCell.getDateCellValue();
			date.setHours(date.getHours() + cancel_Hour_Offset);
			date.setMinutes(date.getMinutes() + cancel_Minute_Offset);
			String day = basicTimeFormat.format(date).toString().substring(0, 8); // just
																					// need
																					// day,ignore
																					// the
																					// hour
																					// and
																					// minute

			String flight = String.valueOf(((int) row.getCell(cancel_Fight_Col).getNumericCellValue()));
			String departure = row.getCell(cancel_Takeoff_col).getStringCellValue().trim();
			String arrival = row.getCell(cancel_Arrival_col).getStringCellValue().trim();

			RowHash rowhash = new RowHash(flight, day, departure, arrival);

			Integer rowid = f_row.get(rowhash);

			/*
			 * if(Debug) System.out.println(" fligth:" + flight + " Date:"+day+
			 * " rowid:"+ rowid);
			 */

			if (rowid != null) {
				// find the row in result table;
				sink = result.getRow(rowid);
				refreshRow(row, sink, resultref.cancellist, cancel_Hour_Offset);
			} else {
				printInfo(InfoLevel.Warn, "Cancel Flight(row: " + (i + 1) + ")" + rowhash.toString()
						+ " doesn't have match row in basic table!");
			}

		}
		printInfo(InfoLevel.Info, "Cancel table:Handle "
				+ (cancel.getLastRowNum() + 1 - cancel_Header_Row - cancel_Footer_Row) + " row in the cancel table");
		// cancelwb.close();

	}

	@SuppressWarnings("deprecation")
	static private void handleSwap1(Sheet swap1, Sheet result, HashMap<RowHash, Integer> f_row, ResultRef resultref)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {

		/*
		 * boolean handle1=true,handle2=true; XSSFWorkbook swapwb=null; Sheet
		 * swap1=null,swap2=null; swapwb=new XSSFWorkbook(OPCPackage.open(new
		 * File(swapname))); swap1=swapwb.getSheetAt(swap_sheet1);
		 * swap2=swapwb.getSheetAt(swap_sheet2);
		 * 
		 */

		if (!swap1.getSheetName().equals(swap_sheet_name1)) {
			printInfo(InfoLevel.Error, "Serious Error:Swap Table:" + swap_sheet1 + " sheet of swap table is not named "
					+ swap_sheet_name1);
			return;
			// handle1=false;
		}

		Row sink;
		// for(int i=Swap_Header_Row;handle1 &&
		// i<=swap1.getLastRowNum()-Swap_Footer_Row;i++)
		for (int i = swap_Header_Row; i <= swap1.getLastRowNum() - swap_Footer_Row; i++) {
			Row row = swap1.getRow(i);

			Cell dateCell = row.getCell(swap_Departure_Col1);
			Cell hourCell = row.getCell(swap_Departure_Hour_Col1);
			if (dateCell.getCellTypeEnum() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(dateCell)) {
				// System.out.println(" Error:Row "+(i+1) + " Col "+
				// Swap_Departure_Col1 +" is not date formate!");
				printInfo(InfoLevel.Warn,
						"Swap first sheet Row " + (i + 1) + " Col " + swap_Departure_Col1 + " is not date formate!");
				continue;

			}

			Date date = dateCell.getDateCellValue();
			int hour = (int) hourCell.getNumericCellValue();
			int minute = hour % 100;
			hour = hour / 100;

			date.setHours(hour + swap_Hour_Offset);
			date.setMinutes(minute + swap_Minute_Offset);

			String day = swapTimeFormat.format(date).toString();

			String flight = String.valueOf(((int) row.getCell(swap_Fight_Col1).getNumericCellValue()));
			String departure = row.getCell(swap_Takeoff_col1).getStringCellValue().trim();
			String arrival = row.getCell(swap_Arrival_col1).getStringCellValue().trim();

			RowHash rowhash = new RowHash(flight, day, departure, arrival);

			Integer rowid = f_row.get(rowhash);

			/*
			 * if(Debug) System.out.println(" fligth:" + flight + " Date:"+day+
			 * " rowid:"+ rowid);
			 */

			if (rowid != null) {
				// find the row in result table;
				sink = result.getRow(rowid);
				// logger.info("--------Swap update:flight(" +
				// rowhash.toString() + " ) rowid:"+ (i+1)+" according first
				// sheet.");
				refreshSwapRow(row, sink, resultref.swaplist, 0, swap_Hour_Offset);
			} else {
				printInfo(InfoLevel.Warn, "Swap Flight(row: " + (i + 1) + ") " + rowhash.toString()
						+ " in the first sheet doesn't have match row in basic table!");
			}

		}
		printInfo(InfoLevel.Info, "Swap table(" + swap_sheet_name1 + "):Handle "
				+ (swap1.getLastRowNum() + 1 - swap_Footer_Row - swap_Header_Row) + " row in the swap table");

		// swapwb.close();

	}

	@SuppressWarnings("deprecation")
	static private void handleSwap2(Sheet swap2, Sheet result, HashMap<RowHash, Integer> f_row, ResultRef resultref)
			throws InvocationTargetException, InterruptedException {
		if (!swap2.getSheetName().equals(swap_sheet_name2)) {
			printInfo(InfoLevel.Error, "Serious Error:Swap Table:" + swap_sheet2 + " sheet of swap table is not named "
					+ swap_sheet_name2);
			// handle2=false;
			return;
		}
		Row sink;
		for (int i = swap_Header_Row; i <= swap2.getLastRowNum() - swap_Footer_Row; i++) {
			Row row = swap2.getRow(i);

			Cell dateCell = row.getCell(swap_Departure_Col2);
			Cell hourCell = row.getCell(swap_Departure_Hour_Col2);
			if (dateCell.getCellTypeEnum() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(dateCell)) {
				printInfo(InfoLevel.Warn,
						"Swap second sheet  Row " + (i + 1) + " Col " + swap_Departure_Col2 + " is not date formate!");
				continue;
			}
			Date date = dateCell.getDateCellValue();
			int hour = (int) hourCell.getNumericCellValue();
			int minute = hour % 100;
			hour = hour / 100;

			date.setHours(hour + swap_Hour_Offset);
			date.setMinutes(minute + swap_Minute_Offset);

			String day = swapTimeFormat.format(date).toString();

			String flight = String.valueOf(((int) row.getCell(swap_Fight_Col2).getNumericCellValue()));
			String departure = row.getCell(swap_Takeoff_col2).getStringCellValue().trim();
			String arrival = row.getCell(swap_Arrival_col2).getStringCellValue().trim();

			RowHash rowhash = new RowHash(flight, day, departure, arrival);

			Integer rowid = f_row.get(rowhash);

			/*
			 * if(Debug) System.out.println(" fligth:" + flight + " Date:"+day+
			 * " rowid:"+ rowid);
			 */

			if (rowid != null) {
				// find the row in result table;
				sink = result.getRow(rowid);
				// printInfo(InfoLevel.Info,"Swap update:flight(" +
				// rowhash.toString() + " ) rowid:"+ (i+1)+" according first
				// sheet.");
				refreshSwapRow(row, sink, resultref.swaplist, 1, swap_Hour_Offset);
			} else {
				printInfo(InfoLevel.Warn, "Swap Flight(row: " + (i + 1) + ") " + rowhash.toString()
						+ " in the second sheet doesn't have match row in basic table!");
			}

		}

		printInfo(InfoLevel.Info, "Swap table(" + swap_sheet_name2 + "):Handle "
				+ (swap2.getLastRowNum() + 1 - swap_Footer_Row - swap_Header_Row) + " row in the swap table");

	}

	static private boolean refreshSwapRow(Row source, Row sink, ArrayList<ColRef> list, int index, int timezoneoffset)
			throws InvocationTargetException, InterruptedException {
		for (int i = 0; i < list.size(); i++) {
			ColRef ref = list.get(i);
			Entry abnormal = ref.abnormal;
			if (abnormal.val != null)// handle fix value;
			{
				if (sink.getCell(ref.resultcol) != null)
					sink.getCell(ref.resultcol).setCellValue(abnormal.val);
				else
					sink.createCell(ref.resultcol).setCellValue(abnormal.val);

			} else if (abnormal.func != null) {// handle abnormal func;

				if (abnormal.func.equals("S-") || abnormal.func.equals("S")) {
					int sourcecol;
					if (index == 0)
						sourcecol = swap_Aircraft_Col1;
					else
						sourcecol = swap_Aircraft_Col2;

					String value = TableTool.getStringValue(source.getCell(sourcecol));

					Cell cell2 = sink.getCell(ref.resultcol);
					if (cell2 == null)
						cell2 = sink.createCell(ref.resultcol);

					if (abnormal.func.equals("S-"))
						value = value.substring(0, 1) + "-" + value.substring(1);

					cell2.setCellValue(value);

				} else
					printInfo(InfoLevel.Warn, "Unknown abnormal func(" + abnormal.func + ") at handle refreshSwapRow:");
			} else if (abnormal.colindex != ResultRef.Blank)// refresh according
															// source cell
			{
				Cell cell1 = source.getCell(abnormal.colindex);

				Cell cell2 = sink.getCell(ref.resultcol);
				if (cell2 == null)
					cell2 = sink.createCell(ref.resultcol);

				cell2.setCellValue(TableTool.getStringValue(cell1));
			} else {
				printInfo(InfoLevel.Warn, "Unknown case at handle refreshRow:");
				ref.println();
			}
		}

		return true;
	}

	static private boolean refreshRow(Row source, Row sink, ArrayList<ColRef> list, int timezoneoffset)
			throws InvocationTargetException, InterruptedException {
		for (int i = 0; i < list.size(); i++) {
			ColRef ref = list.get(i);
			Entry abnormal = ref.abnormal;
			if (abnormal.val != null)///////////////////////// fix value if
										///////////////////////// there is a row
										///////////////////////// in the
										///////////////////////// exception
										///////////////////////// row,then set
										///////////////////////// to the special
										///////////////////////// value,like
										///////////////////////// "D","C";
			{
				if (sink.getCell(ref.resultcol) != null)
					sink.getCell(ref.resultcol).setCellValue(abnormal.val);
				else
					sink.createCell(ref.resultcol).setCellValue(abnormal.val);
			} else if (abnormal.func != null)//////////////////////// function
												//////////////////////// value
			{
				// now don't support any func,cause getFuncValue is nothing
				printInfo(InfoLevel.Warn, "Warn:don't support exception func:" + abnormal.func);
				/*
				 * if(sink.getCell(ref.resultcol)!=null)
				 * sink.getCell(ref.resultcol).setCellValue(
				 * getFuncValue(abnormal,source) ); else
				 * sink.createCell(ref.resultcol).setCellValue(
				 * getFuncValue(abnormal,source) );
				 */
			} else if (abnormal.colindex != ResultRef.Blank) ////////////////// set
																////////////////// value
																////////////////// from
																////////////////// source
																////////////////// cell
			{// set cell value according date format
				Cell cell1 = source.getCell(abnormal.colindex);

				Cell cell2 = sink.getCell(ref.resultcol);
				if (cell2 == null)
					cell2 = sink.createCell(ref.resultcol);

				if (cell1.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell1))
					TableTool.setDateValue(cell1, cell2, ref, timezoneoffset);
				else if (ref.type != null) {
					/*
					 * int mid= ref.type.indexOf('_'); String
					 * type=ref.type.substring(0,mid); String
					 * format=ref.type.substring(mid+1);
					 */
					printInfo(InfoLevel.Error, "Error data type configured in ref file" + ref.type);
				} else
					cell2.setCellValue(TableTool.getStringValue(cell1));

			} else {
				printInfo(InfoLevel.Warn, "Unknown case at handle refreshRow:" + ref.toString());
				ref.println();
			}
		}

		return true;
	}

	static private void init_Airport_Format_data(File file)
			throws InvalidFormatException, IOException, InvocationTargetException, InterruptedException {

		XSSFWorkbook wb;

		wb = new XSSFWorkbook(OPCPackage.open(file));
		Sheet sheet = wb.getSheetAt(0);

		Row row;
		Cell cell1, cell2;
		int i;
		for (i = airportCode_Header_Row; i <= sheet.getLastRowNum() - airportCode_Footer_Row; i++) {
			row = sheet.getRow(i);
			cell1 = row.getCell(airport_format1_col);
			cell2 = row.getCell(airport_format2_col);

			if (!TableTool.NullCell(cell1) && !TableTool.NullCell(cell2)) {
				ReadExcel.airport_Transform.put(cell1.getStringCellValue().trim(), cell2.getStringCellValue().trim());
			} else {
				printInfo(InfoLevel.Warn,
						"AirportCode sheet in AirportCode transform table has empty row" + (i + 1) + "");
			}
		}

		printInfo(InfoLevel.Info,
				"Airport code format file:read "
						+ (sheet.getLastRowNum() + 1 - airportCode_Header_Row - airportCode_Footer_Row)
						+ " row airport map data to program");

		wb.close();
	}

	static private boolean getF_T_Map(XSSFWorkbook wb, int index, String name, HashMap<String, String> f_tail)
			throws InvocationTargetException, InterruptedException {
		Sheet sheet = wb.getSheetAt(index);
		if (!sheet.getSheetName().equals(name)) {
			printInfo(InfoLevel.Error, (index) + " sheet of basic table is not Aircraft_Data");
			return false;
		}

		Cell cell1, cell2;
		Row row;

		for (int i = basic_AircraftData_Header_Row; i <= sheet.getLastRowNum() - basic_AircraftData_Footer_Row; i++) {
			row = sheet.getRow(i);
			cell1 = row.getCell(basicAircraft_TailNumberCol);
			cell2 = row.getCell(basicAircraft_FleetIDCol);

			if (!TableTool.NullCell(cell1) && !TableTool.NullCell(cell2))
				f_tail.put(TableTool.getStringValue(cell1).trim(), TableTool.getStringValue(cell2).trim());
			else
				printInfo(InfoLevel.Warn, "Aircraft sheet in basicdata table has an empty row:" + (i + 1));
		}
		return true;
	}

	static private String getSpecialVal(String hash, String func)
			throws InvocationTargetException, InterruptedException {
		if (func.equals("F_T")) {
			String val = f_tail.get(hash);
			if (val != null) {
				val = "B" + val;
				return val;
			} else {
				printInfo(InfoLevel.Warn, "Dosn't find match FleetID for TailNumber(" + hash + ")");
				return String_Empty;
			}
		} else if (func.equals("MF")) {
			String val = "MF" + hash;
			return val;
		} else if (func.equals("AT")) {
			if (ReadExcel.airport_Format == Airport_Format.IATA)
				return hash;

			String val = ReadExcel.airport_Transform.get(hash);
			if (val != null)
				return val;
			else {
				printInfo(InfoLevel.Warn,
						"Airport code transfer:There is an airport that doesn't have match in hashtabl->" + hash);
				return String_Empty;
			}
		} else {
			printInfo(InfoLevel.Warn, "Unsupport special func :" + func);
			return String_Empty;
		}
	}

	/**
	 * 
	 * @author LancelotWG
	 * @param wb
	 * @param excelname
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws InvocationTargetException
	 * @throws InterruptedException
	 */
	static private void handleBrokenFlight(String abnormalname, ResultRef resultref, String brokenResult)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {
		XSSFWorkbook abnormalwb = new XSSFWorkbook(OPCPackage.open(new File(abnormalname)));
		Sheet cancel = abnormalwb.getSheetAt(cancel_sheet);
		if (cancel != null) {
			if (!cancel.getSheetName().equals(cancel_sheet_name)) {
				printInfo(InfoLevel.Error, "Serious Error: Cancel table" + cancel_sheet
						+ " sheet of cancel table is not named " + cancel_sheet_name);
				return;
			}
			for (int i = cancel_Header_Row; i <= cancel.getLastRowNum() - cancel_Footer_Row; i++) {
				Row row = cancel.getRow(i);
				/**
				 * 增加取消航班的Broken记录
				 */
				String flight = String.valueOf(((int) row.getCell(cancel_Fight_Col).getNumericCellValue()));
				String departure = row.getCell(cancel_Takeoff_col).getStringCellValue().trim();
				String arrival = row.getCell(cancel_Arrival_col).getStringCellValue().trim();
				String tail = row.getCell(cancel_Tail_Col).getStringCellValue().trim();
				String pax = String.valueOf(((int) row.getCell(cancel_Pax_col).getNumericCellValue()));

				Date arrivalTime = row.getCell(cancel_Arrival_Col).getDateCellValue();
				arrivalTime.setHours(arrivalTime.getHours() + cancel_Hour_Offset);
				arrivalTime.setMinutes(arrivalTime.getMinutes() + cancel_Minute_Offset);

				Date departureTime = row.getCell(cancel_Departure_Col).getDateCellValue();
				departureTime.setHours(departureTime.getHours() + cancel_Hour_Offset);
				departureTime.setMinutes(departureTime.getMinutes() + cancel_Minute_Offset);

				for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
					Flights flights = (Flights) iterator.next();
					if (flights.number.flightNumber.equals(flight)) {
						ArrayList<Place> place = flights.places;
						for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
							Place place2 = (Place) iterator2.next();
							if (place2.tailNumber.equals(tail) && place2.arrival.equals(arrival)
									&& place2.departure.equals(departure) && place2.arrivalTime.equals(arrivalTime)
									&& place2.departureTime.equals(departureTime)) {
								place2.setBrokenStatus(BrokenStatus.cancel);
								flights.isBreak = true;
								place2.setPax(pax);
							}
						}

					}
				}
			}
		} else
			printInfo(InfoLevel.Error, "Serous Error:sheet(" + cancel_sheet + ") for cancel data is not exist");
		Sheet delay = abnormalwb.getSheetAt(delay_sheet);
		if (delay != null){
			if (!delay.getSheetName().equals(delay_sheet_name)) {
				printInfo(InfoLevel.Error,
						"Serous Error:delay table:" + (delay_sheet) + " sheet is not named " + delay_sheet_name);
				return;
			}
			for (int i = delay_Header_Row; i <= delay.getLastRowNum() - delay_Footer_Row; i++) {
				Row row = delay.getRow(i);
				/**
				 * 增加延误航班的Broken记录
				 */
				String flight = String.valueOf(((int) row.getCell(delay_Fight_Col).getNumericCellValue()));
				String departure = row.getCell(delay_Takeoff_col).getStringCellValue().trim();
				String arrival = row.getCell(delay_Arrival_col).getStringCellValue().trim();
				String tail = row.getCell(delay_Tail_Col).getStringCellValue().trim();
				String pax = String.valueOf(((int) row.getCell(delay_Pax_col).getNumericCellValue()));

				Date arrivalTime = row.getCell(delay_Arrival_Col).getDateCellValue();
				arrivalTime.setHours(arrivalTime.getHours() + swap_Hour_Offset);
				arrivalTime.setMinutes(arrivalTime.getMinutes() + swap_Minute_Offset);

				Date departureTime = row.getCell(delay_Departure_Col).getDateCellValue();
				departureTime.setHours(departureTime.getHours() + swap_Hour_Offset);
				departureTime.setMinutes(departureTime.getMinutes() + swap_Minute_Offset);

				Date eArrivalTime = row.getCell(delay_EArrival_Col).getDateCellValue();
				eArrivalTime.setHours(eArrivalTime.getHours() + swap_Hour_Offset);
				eArrivalTime.setMinutes(eArrivalTime.getMinutes() + swap_Minute_Offset);

				Date eDepartureTime = row.getCell(delay_EDeparture_Col).getDateCellValue();
				eDepartureTime.setHours(eDepartureTime.getHours() + swap_Hour_Offset);
				eDepartureTime.setMinutes(eDepartureTime.getMinutes() + swap_Minute_Offset);

				for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
					Flights flights = (Flights) iterator.next();
					if (flights.number.flightNumber.equals(flight)) {
						ArrayList<Place> place = flights.places;
						for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
							Place place2 = (Place) iterator2.next();
							if (place2.tailNumber.equals(tail) && place2.arrival.equals(arrival)
									&& place2.departure.equals(departure) && place2.arrivalTime.equals(arrivalTime)
									&& place2.departureTime.equals(departureTime)) {
								place2.setBrokenStatus(BrokenStatus.delay);
								place2.setExpectedDepartureTime(eDepartureTime);
								place2.setExpectedArrivalTime(eArrivalTime);
								place2.setPax(pax);
							}
						}

					}
				}
			}
			for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
				Flights flights = (Flights) iterator.next();
				ArrayList<Place> place = flights.places;
				Date markTime = new Date(0);
				Place forwardPlace = null;
				long timeout = turn_around_time_upperlimit * 60 * 1000;
				long intime = turn_around_time_lowerlimit * 60 * 1000;
				for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
					Place place2 = (Place) iterator2.next();
					if (place2.getBrokenStatus() == BrokenStatus.delay) {
						if (place2.expectedDepartureTime.getTime() - markTime.getTime() <= timeout
								&& place2.expectedDepartureTime.getTime() - markTime.getTime() >= intime) {

						} else {
							if (forwardPlace != null) {
								flights.isBreak = true;
							}
						}
						markTime = place2.getExpectedArrivalTime();
					} else {
						if (place2.departureTime.getTime() - markTime.getTime() <= timeout
								&& place2.departureTime.getTime() - markTime.getTime() >= intime) {

						} else {
							if (forwardPlace != null) {
								flights.isBreak = true;
							}
						}
						markTime = place2.getArrivalTime();
					}
					forwardPlace = place2;
				}

			}
		}else
			printInfo(InfoLevel.Error, "Serous Error:sheet(" + delay_sheet + ") for deley data is not existsheet");
		Sheet swap1 = abnormalwb.getSheetAt(swap_sheet1);
		if (swap1 != null){
			if (!swap1.getSheetName().equals(swap_sheet_name1)) {
				printInfo(InfoLevel.Error, "Serious Error:Swap Table:" + swap_sheet1 + " sheet of swap table is not named "
						+ swap_sheet_name1);
				return;
			}
			for (int i = swap_Header_Row; i <= swap1.getLastRowNum() - swap_Footer_Row; i++) {
				Row row = swap1.getRow(i);
				/**
				 * 增加更换飞机的Broken记录
				 */
				Cell dateCell = row.getCell(swap_Departure_Col1);
				Cell hourCell = row.getCell(swap_Departure_Hour_Col1);
				Date date = dateCell.getDateCellValue();
				int hour = (int) hourCell.getNumericCellValue();
				int minute = hour % 100;
				hour = hour / 100;
				date.setHours(hour + swap_Hour_Offset);
				date.setMinutes(minute + swap_Minute_Offset);
				String flight = String.valueOf(((int) row.getCell(swap_Fight_Col1).getNumericCellValue()));
				String departure = row.getCell(swap_Takeoff_col1).getStringCellValue().trim();
				String arrival = row.getCell(swap_Arrival_col1).getStringCellValue().trim();
				String tail = row.getCell(swap_Tail_Col1).getStringCellValue().trim();
				String pax = String.valueOf(((int) row.getCell(swap_Pax_col1).getNumericCellValue()));
				String newAircraft = row.getCell(swap_Aircraft_Col1).getStringCellValue().trim();
				for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
					Flights flights = (Flights) iterator.next();
					if (flights.number.flightNumber.equals(flight)) {
						ArrayList<Place> place = flights.places;
						for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
							Place place2 = (Place) iterator2.next();
							if (place2.tailNumber.equals(tail) && place2.arrival.equals(arrival)
									&& place2.departure.equals(departure) && date.equals(place2.departureTime)) {
								place2.setBrokenStatus(BrokenStatus.swap);
								place2.setPax(pax);
								place2.setNewFlight(newAircraft);
							}
						}

					}
				}
			}
			for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
				Flights flights = (Flights) iterator.next();
				ArrayList<Place> place = flights.places;
				String aircraft = "";
				boolean isNormal = true;
				if (place.size() >= 2) {
					if (place.get(0).getBrokenStatus() == BrokenStatus.swap) {
						aircraft = place.get(0).getNewFlight();
					} else {
						aircraft = place.get(0).getTailNumber();
					}
				}

				for (Iterator iterator2 = place.iterator(); iterator2.hasNext();) {
					Place place2 = (Place) iterator2.next();
					if (place2.getBrokenStatus() == BrokenStatus.swap) {
						if (!aircraft.equals(place2.newFlight)) {
							isNormal = false;
							break;
						}
					} else {
						if (!aircraft.equals(place2.tailNumber)) {
							isNormal = false;
							break;
						}
					}
				}
				if (!isNormal) {
					/*
					 * for (Iterator iterator2 = place.iterator();
					 * iterator2.hasNext();) { Place place2 = (Place)
					 * iterator2.next(); if (place2.getBrokenStatus() ==
					 * BrokenStatus.swap) {
					 * place2.setBrokenStatus(BrokenStatus.normal); } }
					 */
					flights.isBreak = true;
				}
			}
		}else
			printInfo(InfoLevel.Error,
					"Serous Error:sheet(" + swap_sheet1 + ") for swap data(swaps within subfleets) is not exist");

		/**
		 * 将Broken写入记录类中
		 */
		for (Iterator iterator = f_joint_flights.iterator(); iterator.hasNext();) {
			Flights flights = (Flights) iterator.next();
			/*
			 * ArrayList<Place> place = flights.places; boolean isNormal = true;
			 * for (Iterator iterator2 = place.iterator(); iterator2.hasNext();)
			 * { Place place2 = (Place) iterator2.next(); BrokenStatus status =
			 * place2.getBrokenStatus(); if (status != BrokenStatus.normal) {
			 * isNormal = false; break; } }
			 */
			if (flights.isBreak) {
				broken_records.add(new BrokenRecords(flights.number, flights.places));
			}
		}	
	}
	
	static private void createBrokenFlightResultSheet(String brokenResult)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet brokenResultSheet = wb.createSheet(broken_result_sheet_name);
		// copy the header
		printInfo(InfoLevel.Info, "checked the result table:" + brokenResult);
		File file = new File(brokenResult);
		if (!file.exists()) {
			printInfo(InfoLevel.Error, "Serious Error:Result file The Output file(" + brokenResult + ") is not exist!");
			return;
		}
		XSSFWorkbook resultwb = new XSSFWorkbook(OPCPackage.open(file));
		Sheet result = resultwb.getSheetAt(broken_result_sheet);
		Row row = result.getRow(0);
		Row sink = brokenResultSheet.createRow(0);
		sink.createCell(0).setCellValue(row.getCell(0).getStringCellValue().trim());
		
		sink.createCell(1).setCellValue(broken_records.size());
		row = result.getRow(1);
		sink = brokenResultSheet.createRow(1);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			// printInfo(InfoLevel.Info, "-> " + i + );
			sink.createCell(i).setCellValue(row.getCell(i).getStringCellValue().trim());
		}
		Cell cell;
		Row sink1;
		int index = broken_Header_Row;
		SimpleDateFormat dateformat = new SimpleDateFormat("yyyy-MM-dd HH:mm:SS", Locale.US);
		for (int i = 0; i <= broken_records.size() - 1; i++) {
			BrokenRecords brokenRecords = broken_records.get(i);
			ArrayList<Place> places = brokenRecords.places;
			for (Iterator iterator = places.iterator(); iterator.hasNext();) {
				Place place = (Place) iterator.next();
				sink1 = brokenResultSheet.createRow(index);
				cell = sink1.createCell(0);
				cell.setCellValue(brokenRecords.number.flightNumber);
				cell = sink1.createCell(1);
				cell.setCellValue(dateformat.format(place.departureTime));
				cell = sink1.createCell(2);
				cell.setCellValue(dateformat.format(place.arrivalTime));
				cell = sink1.createCell(3);
				cell.setCellValue(place.departure);
				cell = sink1.createCell(4);
				cell.setCellValue(place.arrival);
				cell = sink1.createCell(5);
				cell.setCellValue(place.tailNumber);
				cell = sink1.createCell(6);
				cell.setCellValue(place.brokenStatus.toString());
				cell = sink1.createCell(7);
				if (place.newFlight == null) {
					cell.setCellValue("");
				} else {
					cell.setCellValue(place.newFlight);
				}
				cell = sink1.createCell(8);
				if (place.expectedDepartureTime != null) {
					cell.setCellValue(dateformat.format(place.expectedDepartureTime));
				} else {
					cell.setCellValue("");
				}
				cell = sink1.createCell(9);
				if (place.expectedArrivalTime != null) {
					cell.setCellValue(dateformat.format(place.expectedArrivalTime));
				} else {
					cell.setCellValue("");
				}
				cell = sink1.createCell(10);
				cell.setCellValue(place.pax);
				index++;
			}

		}
		resultwb.close();
		// flush the data to file
		FileOutputStream fileOut = new FileOutputStream(brokenResult);
		if (wb != null) {
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} else {
			printInfo(InfoLevel.Error, "Serious Error: flush To File,the workbook is null");
		}
	}

	static private void flushToFile(XSSFWorkbook wb, Sheet sheet, String excelname, String brokenResult)
			throws IOException, InvalidFormatException, InvocationTargetException, InterruptedException {
		if (sheet == null)
			return;

		// copy the header
		printInfo(InfoLevel.Info, "checked the result table:" + excelname);
		File file = new File(excelname);
		if (!file.exists()) {
			printInfo(InfoLevel.Error, "Serious Error:Result file The Output file(" + excelname + ") is not exist!");
			return;
		}

		XSSFWorkbook resultwb = new XSSFWorkbook(OPCPackage.open(file));
		Sheet result = resultwb.getSheetAt(result_sheet);
		Row row = result.getRow(0);

		Row sink = sheet.createRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			// printInfo(InfoLevel.Info, "-> " + i + );
			sink.createCell(i).setCellValue(row.getCell(i).getStringCellValue().trim());
		}

		resultwb.close();
		createBrokenFlightResultSheet(brokenResult);
		printInfo(InfoLevel.Info,
				"Finished processing all input file,and refresh date to the result file:" + excelname);
		
		// flush the data to file
		FileOutputStream fileOut = new FileOutputStream(excelname);
		if (wb != null) {
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} else {
			printInfo(InfoLevel.Error, "Serious Error: flush To File,the workbook is null");
		}

	}

	enum InfoLevel {
		Info, Warn, Error;
		public String getName() {
			return this.name() + ":";
		}
	}

	static void printInfo(InfoLevel level, String info) throws InvocationTargetException, InterruptedException {
		if (level == InfoLevel.Info)
			logger.info(info);
		else if (level == InfoLevel.Warn)
			logger.warn(info);
		else
			logger.error(info);

		String val = level.getName() + info;
		if (!Main.gui) {
			System.out.println(val);
		} else {
			SwingUtilities.invokeAndWait(new Runnable() {
				public void run() {
					if (level == InfoLevel.Info) {
						((ShowPanel) Main.frame).infota.append(val + "\t\r\n");
					} else if (level == InfoLevel.Warn) {
						((ShowPanel) Main.frame).errorta.append(val + "\t\r\n");
					} else if (level == InfoLevel.Error) {
						((ShowPanel) Main.frame).errorta.append(val + "\t\r\n");
					}

					Main.frame.invalidate();
				}
			});
		}
	}

}