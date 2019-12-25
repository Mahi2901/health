package excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;

public class CellReader
{
	String				fileNameWithPath	= null;
	//FileInputStream		fileInputStream		= null;
	FileOutputStream	fileOutputStream	= null;
	Workbook			workbook			= null;
	Sheet				sheet				= null;
	Row					row;
	Cell				cell;
	int					lastRowNum;
	int					numberOfSheets;
	public int			currentRowIndex		= 0;
	public boolean		trimStringData		= false;

	protected Cell getCell()
	{
		return cell;
	}
	public CellAddress getCellAddress()
	{
		if (cell != null)
		{
			return cell.getAddress();
		}
		else
			return null;
	}
	public int getLastRowNum()
	{
		return lastRowNum;
	}

	public int getNumberOfSheets()
	{
		return numberOfSheets;
	}

	public int getCurrentRowIndex()
	{
		return currentRowIndex;
	}

	public int getCurrentColumnIndex()
	{
		return currentColumnIndex;
	}
	public int			currentColumnIndex	= 0;
	private CellData	cellData			= null;

	public void newCellData()
	{
		cellData = new CellData();
	}

	public void setCellData(CellData cellData)
	{
		this.cellData = cellData;
	}

	public void setDefault()
	{
		cellData.setDefault();
	}

	public CellReader(String fileNameWithPath) throws Exception
	{
		try
		{
			this.fileNameWithPath = fileNameWithPath;
			File file = new File(this.fileNameWithPath);
			String extention = FilenameUtils.getExtension(file.getName());
			if (extention != null)
			{
				extention = extention.toUpperCase().trim();
			}
			else
			{
				extention = "";
			}
			//fileInputStream = new FileInputStream(fileNameWithPath);
			//this.workbook = WorkbookFactory.create(fileInputStream);
			if (file.exists())
			{
				//following not work for few file
				//this.workbook = WorkbookFactory.create(file);

				//work for all file
				FileInputStream fileInputStream = new FileInputStream(fileNameWithPath);
				this.workbook = WorkbookFactory.create(fileInputStream);
			}
			else
			{
				if (extention.equalsIgnoreCase("xlsx"))
				{
					this.workbook = new XSSFWorkbook();
				}
				else
				{
					this.workbook = new HSSFWorkbook();
				}
			}
			this.numberOfSheets = this.workbook.getNumberOfSheets();
			newCellData();
		}
		catch (InvalidFormatException invalidFormatException)
		{
			invalidFormatException.printStackTrace();
			throw new Exception("invalid Format Exception" + invalidFormatException.getMessage());
		}
	}
	// public void
	public void setSheet(Sheet sheet)
	{
		this.sheet = sheet;
		int rowcount = sheet.getLastRowNum();
	}
	public void setAutoColumnWidth()
	{
		//please do not change liberary code
		//		int myLastRowNum = sheet.getLastRowNum() + 2;
		//		for (int i = 0; i < myLastRowNum; i++)
		//		{
		//			this.sheet.autoSizeColumn(i);
		//		}
		if (row != null)
		{
			for (int i = 0; i < row.getLastCellNum(); i++)
			{
				System.out.println("set-AutoWidth at i=" + i);
				this.sheet.autoSizeColumn(i);
			}
		}
	}

	private static void frame(CellRangeAddress region, Sheet sheet, Workbook wb)
	{
		sheet.addMergedRegion(region);

		final short borderMediumDashed = CellStyle.BORDER_THIN;
		RegionUtil.setBorderBottom(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderTop(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderLeft(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderRight(borderMediumDashed, region, sheet, wb);
	}
	public void setmergeColumn()
	{
		//sheet.addMergedRegion(rowFrom,rowTo,colFrom,colTo);
		frame(new CellRangeAddress(1, 1, 1, 5), this.sheet, this.workbook);
		frame(new CellRangeAddress(2, 2, 1, 5), this.sheet, this.workbook);
		frame(new CellRangeAddress(3, 3, 1, 5), this.sheet, this.workbook);
		frame(new CellRangeAddress(4, 4, 1, 5), this.sheet, this.workbook);
		frame(new CellRangeAddress(1, 1, 12, 14), this.sheet, this.workbook);

	}
	public void setmergeColumn_Admin_Preview()
	{
		//sheet.addMergedRegion(rowFrom,rowTo,colFrom,colTo);
		frame(new CellRangeAddress(2, 2, 0, 1), this.sheet, this.workbook);
		frame(new CellRangeAddress(3, 3, 0, 1), this.sheet, this.workbook);

	}
	public void setmergeColumn_Admin_Preview1(int rowFrom, int rowTo)
	{
		//sheet.addMergedRegion(rowFrom,rowTo,colFrom,colTo);
		frame(new CellRangeAddress(rowFrom, rowTo, 0, 1), this.sheet, this.workbook);
	}
	public void setmergeColumn_Admin_Preview2(int rowFrom, int rowTo, int colFrom, int colTo)
	{
		//sheet.addMergedRegion(rowFrom,rowTo,colFrom,colTo);
		frame(new CellRangeAddress(rowFrom, rowTo, colFrom, colTo), this.sheet, this.workbook);
	}
	public void setmergeColumnVia(int firstCellraw, int lastCellraw, int firstCellcol, int lastCellcol)
	{

		frame(new CellRangeAddress(firstCellraw, lastCellraw, firstCellcol, lastCellcol), this.sheet, this.workbook);

	}
	public void setAutoColumnWidth1_contract()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(0, 5000);
			sheet.setColumnWidth(1, 6000);
			sheet.setColumnWidth(2, 4000);
			sheet.setColumnWidth(3, 8000);
			sheet.setColumnWidth(4, 4500);
			sheet.setColumnWidth(5, 4500);
			sheet.setColumnWidth(6, 5000);
			sheet.setColumnWidth(7, 5000);

			sheet.setColumnWidth(8, 5000);
			sheet.setColumnWidth(9, 5000);
			sheet.setColumnWidth(10, 4000);
			sheet.setColumnWidth(11, 8000);
			sheet.setColumnWidth(12, 8000);
			sheet.setColumnWidth(13, 5000);
			sheet.setColumnWidth(14, 5000);
			sheet.setColumnWidth(15, 8000);
			sheet.setColumnWidth(16, 7000);
			sheet.setColumnWidth(17, 5000);
			sheet.setColumnWidth(18, 8000);
			sheet.setColumnWidth(19, 6000);
			sheet.setColumnWidth(20, 6000);
			sheet.setColumnWidth(21, 15000);

		}

	}
	public void setAutoColumnWidthcontract3(int raw)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(raw, 5000);
		}

	}
	public void setAutoColumnWidthcontract2(int raw)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(raw, 10000);
		}
	}
	public void setCellRangeAddress(int firstCellraw, int lastCellraw, int firstCellcol, int lastCellcol)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;

		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setAutoFilter(new CellRangeAddress(firstCellraw, lastCellraw, firstCellcol, lastCellcol));
		}

	}
	public void setAutoColumnWidthdefault3(int raw)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(raw, 6000);
		}
	}
	public void setAutoColumnWidth_contract1()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(5, 10000);
		}

	}
	public void setAutoColumnWidth2()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(12, 10000);

		}

	}
	public void setAutoColumnWidth3(int raw)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(raw, 10000);

		}

	}
	public void setAutoColumnWidth_VIA(int raw)
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(raw, 5000);
		}

	}
	public void setColumnWidth_14_Type()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{

			sheet.setColumnWidth(0, 9000);
			sheet.setColumnWidth(1, 5000);
			sheet.setColumnWidth(2, 5000);
			sheet.setColumnWidth(3, 2500);
			sheet.setColumnWidth(4, 2500);
			sheet.setColumnWidth(5, 2500);
			sheet.setColumnWidth(6, 4000);
			sheet.setColumnWidth(7, 4000);
			sheet.setColumnWidth(8, 4000);
			sheet.setColumnWidth(9, 4000);
			sheet.setColumnWidth(10, 4000);
			sheet.setColumnWidth(11, 5000);
			sheet.setColumnWidth(11, 4000);
			sheet.setColumnWidth(12, 9000);

		}
	}
	public void setSheet(int index) throws Exception
	{
		try
		{
			this.sheet = this.workbook.getSheetAt(index);
		}
		catch (IllegalArgumentException illegalArgumentException)
		{
			throw new Exception("this.workbook.getSheetAt(" + index + "), Sheet not found on given index : err:" + illegalArgumentException.getMessage());//"The file do not contain specified sheet. Please provide the correct sheet number of the usage data."
		}
		lastRowNum = sheet.getLastRowNum();
	}

	public void setSheet(String sheetName) throws Exception
	{
		try
		{
			this.sheet = this.workbook.getSheet(sheetName);

		}
		catch (NullPointerException nullPointerException)
		{
			throw nullPointerException;
		}
		catch (Exception exception)
		{
			throw new Exception("this.workbook.getSheet(" + sheetName + "), Sheet not found on given sheetName");//"The file do not contain specified sheet. Please provide the correct sheet number of the usage data."
		}

		lastRowNum = sheet.getLastRowNum();
	}

	public int[] getCellIndex()
	{
		return new int[] { this.currentRowIndex, this.currentColumnIndex };
	}
	public void bookmark(String bookmarkName, BookmarkCellPointer bookmarkCellPointer)
	{
		bookmarkCellPointer.mark(bookmarkName, getCellIndex());
	}
	public void setCellAddress(String cellAddress)
	{
		// this formula should not likely use
		CellReference cellReference = new CellReference(cellAddress);
		setCellAddress(cellReference.getRow(), cellReference.getCol());
	}
	public void setCellAddress_read(String cellAddress)
	{
		// this formula should not likely use
		CellReference cellReference = new CellReference(cellAddress);
		row = sheet.getRow(cellReference.getRow());
		cell = row.getCell(cellReference.getCol());
		setCellAddress(row.getRowNum(), cell.getColumnIndex());
	}
	public void setCellAddress_read(int rowIndex, int colIndex)
	{
		// this formula should not likely use

		row = sheet.getRow(rowIndex);

		if (row == null)
		{

		}
		else
		{

			cell = row.getCell(colIndex);
			if (cell == null)
			{

				//cell=row.getCell(cell.getColumnIndex(), Row.RETURN_BLANK_AS_NULL);
			}
			else
			{

				setCellAddress(row.getRowNum(), cell.getColumnIndex());
			}

		}

	}
	public int getCellNumericValue(int rowIndex, int colIndex)
	{
		row = sheet.getRow(rowIndex);
		cell = row.getCell(colIndex);
		int Val = (int) cell.getNumericCellValue();
		return Val;
	}
	public double getCellLongValue(int rowIndex, int colIndex)
	{
		row = sheet.getRow(rowIndex);
		cell = row.getCell(colIndex);
		if (cell == null)
			return 0;
		else
			return cell.getNumericCellValue();

	}
	public String getCellStringValue(int rowIndex, int colIndex)
	{
		row = sheet.getRow(rowIndex);
		cell = row.getCell(colIndex);
		if (cell == null)
			return "";
		else
			return cell.getStringCellValue();

	}
	public Date getCellDateValue(int rowIndex, int colIndex)
	{
		row = sheet.getRow(rowIndex);
		cell = row.getCell(colIndex);
		if (cell == null)
			return null;
		else
			return cell.getDateCellValue();

	}
	public void setCellAddress(String bookmarkName, BookmarkCellPointer bookmarkCellPointer)
	{
		int[] tmp = bookmarkCellPointer.get(bookmarkName);
		setCellAddress(tmp[0], tmp[1]);
	}
	public void setCellAddress(ConditionMatcher conditionMatcher)
	{
		setCellAddress(conditionMatcher.currentRowIndex, conditionMatcher.currentColumnIndex);
	}
	public void setCellAddress(int rowIndex, int columnIndex)
	{
		if (this.currentRowIndex != rowIndex)
		{
			this.currentRowIndex = rowIndex;
			changeRow();
		}
		if (this.currentColumnIndex != columnIndex)
		{
			this.currentColumnIndex = columnIndex;
		}
		changeCell();// change enven both are not change
	}

	private void changeRow()
	{
		this.row = sheet.getRow(this.currentRowIndex);
	}

	private void changeCell()
	{
		if (this.row != null)
		{
			this.cell = this.row.getCell(this.currentColumnIndex);
		}
		else
		{
			this.cell = null;
		}
	}

	public void move(int bottom, int right)
	{
		if (bottom != 0)
		{
			currentRowIndex = currentRowIndex + bottom;
			changeRow();
		}
		if (right != 0)
		{
			currentColumnIndex = currentColumnIndex + right;
		}
		changeCell();
	}

	public CellData read()
	{
		return read((byte) 0);
	}
	public CellData read(byte expectedType)
	{
		// Cell cel = sheet.getRow(i).getCell(j);
		if (cell != null)
		{
			cellData = getCellData(cellData, cell, cell.getCellTypeEnum(), expectedType);
			cellData.isCellNull = false;
		}
		else
		{
			cellData.setDefault();
			cellData.isCellNull = true;
			cellData.haveValue = false;
		}
		return cellData;
	}
	private CellData getCellData(CellData cellData, Cell cell, CellType enums, int expectedType)
	{
		cellData.setDefault();
		cellData.errorCellValue = 0;
		cellData.isBlank = false;
		cellData.isNone = false;
		cellData.haveValue = false;

		// int cel_Type = cell.getCellType();
		//cellData.cellType = cell.getCellType();// cell.getCellTypeEnum();
		cellData.currentColumnIndex = this.currentColumnIndex;
		cellData.currentRowIndex = this.currentRowIndex;
		cellData.cellReference = cell.getAddress().toString();
		cellData.debugPoint = "index(" + this.currentRowIndex + "," + this.currentColumnIndex + "),count(" + (this.currentRowIndex + 1) + "," + (this.currentColumnIndex + 1) + ")" + cell.getAddress().toString();
		cellData.cellType = cell.getCellTypeEnum();
		if (cellData.cellType == CellType.FORMULA)
		{
			cellData.cellType = cell.getCachedFormulaResultTypeEnum();
		}
		switch (enums)
		{
			case _NONE:
				cellData.isNone = true;
				break;
			case BLANK:
				cellData.isBlank = true;
				break;
			case BOOLEAN:
				cellData.booleanData = cell.getBooleanCellValue();
				if (cellData.booleanData != null)
				{
					cellData.haveValue = true;
				}
				break;
			case ERROR:
				cellData.errorCellValue = cell.getErrorCellValue();
				if (cellData.errorCellValue != 0)
				{
					cellData.haveValue = true;
				}
				break;
			case FORMULA:
				cellData.formulaString = cell.getCellFormula();
				//FormulaEvaluator evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
				//cellData = getCellData(cellData, cell, evaluator.evaluateFormulaCellEnum(cell), expectedType);
				cellData = getCellData(cellData, cell, cell.getCachedFormulaResultTypeEnum(), expectedType);
				break;
			case NUMERIC:
				cellData.numeric = cell.getNumericCellValue();
				cellData.haveValue = true;
				//Date javaDate = DateUtil.getJavaDate(cellData.numeric);
				break;
			case STRING:
				cellData.stringData = cell.getStringCellValue();
				if (cellData.stringData != null)
				{
					if (trimStringData)
					{
						cellData.stringData = cellData.stringData.trim();
						if (!cellData.stringData.equals(""))
						{
							cellData.haveValue = true;
						}
					}
					else
					{
						cellData.haveValue = true;
					}
				}
				break;
			default:
				cellData.haveValue = false;
				System.out.print("inside the default..Cellread :: case-190");
		}
		return cellData;
	}
	public CellData executeRead()
	{
		return executeRead((byte) 0);
	}
	public CellData executeRead(byte expectedType)
	{
		// Cell cel = sheet.getRow(i).getCell(j);
		if (cell != null)
		{
			cellData = getCellDataAfterExecuteFormula(cellData, cell, cell.getCellTypeEnum(), expectedType);
			cellData.isCellNull = false;
		}
		else
		{
			cellData.setDefault();
			cellData.isCellNull = true;
			cellData.haveValue = false;
		}
		return cellData;
	}
	private CellData getCellDataAfterExecuteFormula(CellData cellData, Cell cell, CellType enums, int expectedType)
	{
		cellData.setDefault();
		cellData.errorCellValue = 0;
		cellData.isBlank = false;
		cellData.isNone = false;
		cellData.haveValue = false;

		// int cel_Type = cell.getCellType();
		//cellData.cellType = cell.getCellType();// cell.getCellTypeEnum();
		cellData.debugPoint = "index(" + this.currentRowIndex + "," + this.currentColumnIndex + "),count(" + (this.currentRowIndex + 1) + "," + (this.currentColumnIndex + 1) + ")" + cell.getAddress().toString();
		cellData.cellType = cell.getCellTypeEnum();
		if (cellData.cellType == CellType.FORMULA)
		{
			//cellData.cellType = cell.getCachedFormulaResultTypeEnum();
			cellData.cellType = cell.getCellTypeEnum();
		}
		switch (enums)
		{
			case _NONE:
				cellData.isNone = true;
				break;
			case BLANK:
				cellData.isBlank = true;
				break;
			case BOOLEAN:
				cellData.booleanData = cell.getBooleanCellValue();
				if (cellData.booleanData != null)
				{
					cellData.haveValue = true;
				}
				break;
			case ERROR:
				cellData.errorCellValue = cell.getErrorCellValue();
				if (cellData.errorCellValue != 0)
				{
					cellData.haveValue = true;
				}
				break;
			case FORMULA:
				cellData.formulaString = cell.getCellFormula();
				//FormulaEvaluator evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
				//cellData = getCellData(cellData, cell, evaluator.evaluateFormulaCellEnum(cell), expectedType);
				//cellData = getCellData(cellData, cell, cell.getCachedFormulaResultTypeEnum(), expectedType);
				FormulaEvaluator evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
				cellData = getCellData(cellData, cell, evaluator.evaluateFormulaCellEnum(cell), expectedType);
				break;
			case NUMERIC:
				cellData.numeric = cell.getNumericCellValue();
				cellData.haveValue = true;
				//Date javaDate = DateUtil.getJavaDate(cellData.numeric);
				break;
			case STRING:
				cellData.stringData = cell.getStringCellValue();
				if (cellData.stringData != null)
				{
					cellData.haveValue = true;
				}
				break;
			default:
				cellData.haveValue = false;
				System.out.print("inside the default..Cellread :: case-190");
		}
		return cellData;
	}
	public void close() throws IOException
	{
		//		if (fileInputStream != null)
		//			fileInputStream.close();
		if (workbook != null)
		{
			workbook.close();
			workbook = null;
		}

	}
	public void setAutoColumnWidthForTemplate()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			this.sheet.autoSizeColumn(1);
			this.sheet.setColumnWidth(2, 5000);
			this.sheet.setColumnWidth(i + 3, 3500);
		}
		//		if (row != null)
		//		{
		//			for (int i = 0; i < row.getLastCellNum(); i++)
		//			{
		//				this.sheet.autoSizeColumn(i);
		//			}
		//		}	
	}
	private static void frameWithOutBorder(CellRangeAddress region, Sheet sheet, Workbook wb)
	{
		sheet.addMergedRegion(region);

		final short borderMediumDashed = CellStyle.BORDER_NONE;
		RegionUtil.setBorderBottom(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderTop(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderLeft(borderMediumDashed, region, sheet, wb);
		RegionUtil.setBorderRight(borderMediumDashed, region, sheet, wb);
	}
	public void setmergeColumnForTemplateHeading(int rowFrom, int rowTo, int colFrom, int colTo)
	{
		//sheet.addMergedRegion(rowFrom,rowTo,colFrom,colTo);
		frameWithOutBorder(new CellRangeAddress(rowFrom, rowTo, colFrom, colTo), this.sheet, this.workbook);
	}
	public void setAutoColumnWidth_AICP()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			sheet.setColumnWidth(0, 10000);
			sheet.setColumnWidth(1, 5000);
			sheet.setColumnWidth(2, 5000);
			sheet.setColumnWidth(3, 5000);
			sheet.setColumnWidth(4, 5000);
			sheet.setColumnWidth(5, 5000);
			sheet.setColumnWidth(6, 5000);
			sheet.setColumnWidth(7, 5000);
			sheet.setColumnWidth(8, 5000);
			sheet.setColumnWidth(9, 5000);
			sheet.setColumnWidth(10, 5000);
			sheet.setColumnWidth(i + 11, 3500);
		}

	}
	public void setAutoColumnWidthForTemplate5()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{
			this.sheet.autoSizeColumn(1);
			this.sheet.setColumnWidth(2, 4000);
			this.sheet.setColumnWidth(i + 3, 3500);
		}
		//		if (row != null)
		//		{
		//			for (int i = 0; i < row.getLastCellNum(); i++)
		//			{
		//				this.sheet.autoSizeColumn(i);
		//			}
		//		}	
	}

	public void setAutoColumnWidthForTemplatePhotoGraphyAndCGI()
	{
		int myLastRowNum = sheet.getLastRowNum() + 2;
		for (int i = 0; i < myLastRowNum; i++)
		{

			if (i == 0)
			{

			}
			else
			{
				this.sheet.autoSizeColumn(1);
				this.sheet.setColumnWidth(2, 3780);
				this.sheet.setColumnWidth(3, 3780);
				this.sheet.setColumnWidth(4, 3780);
				this.sheet.setColumnWidth(5, 3780);

				this.sheet.setColumnWidth(i * 4 + 2, 3780);
				this.sheet.setColumnWidth(i * 4 + 3, 3780);
				this.sheet.setColumnWidth(i * 4 + 4, 3780);
				this.sheet.setColumnWidth(i * 4 + 5, 3780);
			}
			/*sheet.setColumnWidth(2, 4000);
			sheet.setColumnWidth(3, 4000);
			sheet.setColumnWidth(4, 4000);
			sheet.setColumnWidth(5, 4000);
			sheet.setColumnWidth(6, 4000);
			sheet.setColumnWidth(7, 4000);
			sheet.setColumnWidth(8, 4000);
			sheet.setColumnWidth(9, 4000);
			sheet.setColumnWidth(10, 4000);*/
			sheet.setColumnWidth(i * 4 + 2, 4200);
			/*	sheet.setColumnWidth(i+11+1,4200);*/
		}
		//		if (row != null)
		//		{
		//			for (int i = 0; i < row.getLastCellNum(); i++)
		//			{
		//				this.sheet.autoSizeColumn(i);
		//			}
		//		}	
	}
	
	public void setCurrencyFormat(String symbol)
	{
		HSSFWorkbook workbook = new HSSFWorkbook(); 
		//CreationHelper createHelper = workbook.getCreationHelper();
		HSSFCellStyle styleCurrencyFormat = workbook.createCellStyle();
		HSSFDataFormat df = workbook.createDataFormat();
		//styleCurrencyFormat.setDataFormat(createHelper.createDataFormat().getFormat("$#,##0.00"));
		styleCurrencyFormat.setDataFormat(df.getFormat(symbol));
		changeCell();// change enven both are not change
	}
}
