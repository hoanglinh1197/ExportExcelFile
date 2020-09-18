import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Export {
	private XSSFWorkbook workbook;
	private List<XSSFSheet> sheets;
	private String shifts[];
	private String employees[];
	private String employees_shifts[];
	private String schemaName;
	private String destination;

	// path: vị trị của file excel, schemaName : tên schema trong database
	// destination là nơi xuất ra file excel (.sql)
	public Export(String path, String schemaName, String destination) throws IOException {
		workbook = new XSSFWorkbook(path);
		getSheets();
		initTitle();
		this.schemaName = schemaName;
		this.destination = destination;
	}

	public void initTitle() {
		shifts = new String[] { "dayOfWeek", "finishHour", "isDisable", "shiftCode", "shiftCoefficient", "shiftName",
				"startHour" };

		employees = new String[] { "eId", "eAddress", "eEndDate", "eCode", "eFirstName", "isDisable", "eLasttName",
				"ePhoneNumber", "eStartDate", "department_id", "position_id" };

		employees_shifts = new String[] { "emshift_id", "fromDate", "isDisable", "toDate", "employee_id", "shift_id" };

	}
	// Lấy ra danh sách các sheet từ file excel
	private void getSheets() {
		sheets = new ArrayList<XSSFSheet>();
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheets.add(workbook.getSheetAt(i));
		}
	}

	// Trả dữ liệu từ ô trong excel
	public String getStringFromCell(Cell cell, int i, String name) {
		String str = "";
		CellType cellType = cell.getCellType();

		switch (cellType) {
		case _NONE:
			break;
		case BOOLEAN:
			str += cell.getBooleanCellValue();

			break;
		case BLANK:
			str += "\t";
			break;
		case FORMULA:
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			str += evaluator.evaluate(cell).getNumberValue();
			break;
		case NUMERIC:
			str += ((i == 4 || i == 5) && name.equals("employee_shift"))
					? new SimpleDateFormat("yyyy/MM/dd").format(DateUtil.getJavaDate(cell.getNumericCellValue()))
					: Double.valueOf(cell.getNumericCellValue()).intValue();
			break;
		case STRING:
			str += (name.equals("Shift") && (i == 3)) ? fragment(cell.getStringCellValue()) : cell.getStringCellValue();
			break;
		case ERROR:
			str += "error";
			break;
		}
		return str;
	}

	// Trả về từng dòng dlieu trong Sheet
	public String[] getRowInDatasheet(int indexofSheet) {
		List<String> dataInRows = new ArrayList<String>();
		XSSFSheet sheet = sheets.get(indexofSheet);
		Iterator<Row> rows = sheet.iterator();
		String sheetName = sheet.getSheetName();

		while (rows.hasNext()) {
			Row row = rows.next();
			if (row.getRowNum() == 0)
				continue;

			Iterator<Cell> cells = row.cellIterator();
			StringBuffer str = new StringBuffer();
			int i = 0;
			str.append(getDataInRow(cells, sheetName, row, i));
			dataInRows.add(str.toString());
		}
		String[] strs = new String[dataInRows.size()];
		dataInRows.toArray(strs);
		return strs;
	}

	// Định dạng lại dữ liệu theo đúng trường trong database
	public String[] format(String[] row, String sheetName) {
		String[] result = null;

		if (sheetName.equals("Shift")) {
			result = new String[] { null, row[4], null, row[1], null, row[2], row[3] };

		} else if (sheetName.equals("employee")) {
			result = new String[] { row[0], null, null, row[1], row[3], null, row[2], null, null, null, null };

		} else if (sheetName.equals("employee_shift")) {
			result = new String[] { row[0], row[3], null, row[4], row[1], row[2] };
		}
		return result;
	}

	// Trả về câu lệnh insert
	public String getInsertQuery(String[] data, String sheetName) {
		String query = "";
		if (sheetName.equals("Shift")) {
			InsertedQuery inserting = new InsertedQuery(schemaName, "shifts", shifts, data);
			query = inserting.getQuery();
		} else if (sheetName.equals("employee")) {
			InsertedQuery inserting = new InsertedQuery(schemaName, "employees", employees, data);
			query = inserting.getQuery();
		} else if (sheetName.equals("employee_shift")) {
			InsertedQuery inserting = new InsertedQuery(schemaName, "employeeshifts", employees_shifts, data);
			query = inserting.getQuery();
		}
		return query;
	}
	
	// Trả về dsach các query
	public List<String> getQueries(int index) {
		List<String> list = new ArrayList<String>();
		String[] rows = getRowInDatasheet(index);
		String sheetName = sheets.get(index).getSheetName();
		for (String s : rows) {
			String[] data = s.split("\t");
			data = format(data, sheetName);
			list.add(getInsertQuery(data, sheetName));

		}
		return list;
	}

	public String getDataInRow(Iterator<Cell> cells, String sheetName, Row row, int i) {
		StringBuilder str = new StringBuilder();
		while (cells.hasNext()) {
			Cell cell = cells.next();
			i++;
			String data = getStringFromCell(cell, i, sheetName);
			if (i == 3 && (row.getRowNum() != 0) && sheetName.equals("Shift")) {
				String[] names = data.split("\t");
				for (String s : names)
					str.append(s + "\t");
				break;
			} else if (cells.hasNext()) {
				str.append(data + "\t");
			} else {
				str.append(data + "\t");
			}
		}
		return str.substring(0, str.length() - 1);
	}

	// Xử lí cột Name trong bảng Shift(excel): p/tach name thành 3 phần: name,
	// startDate, endDate
	public String fragment(String str) {
		StringBuffer result = new StringBuffer();

		int from = str.indexOf("(");
		result.append(str.substring(0, from) + "\t");

		String[] times = str.substring(from + 1, str.indexOf(")")).split("-");
		result.append(times[0] + "\t" + times[1]);
		return result.toString();
	}
	
	// Xuất ra file excel
	public void export() throws IOException {
		OutputStreamWriter writer = new OutputStreamWriter(new FileOutputStream(destination));
		for (int i = 0; i < sheets.size(); i++) {
			for (String s : getQueries(i)) {
				writer.write(s + "\n");
			}
		}
		System.out.println("Export thành công");
		writer.close();
	}

	public static void main(String[] args) throws IOException {
		Export ex = new Export("D:/Data/Download/data.xlsx", "timekeeper","d:data.sql");
		ex.export();

	}

}
