package excelreader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.util.CellRangeAddress;


public class ExcelReaderWriter extends ExcelReader
{
	ExcelStyleSet excelStyleSet = null;
	public ExcelReaderWriter(String fileNameWithPath) throws Exception
	{
		super(fileNameWithPath);
	}
	public ExcelReaderWriter(String fileNameWithPath, ExcelStyleCreator excelStyleCreator) throws Exception
	{
		super(fileNameWithPath);
		excelStyleSet = new ExcelStyleSet();
		excelStyleCreator.addExcelStyleSet(workbook, excelStyleSet);
		//		ExcelDefaultStyle excelDefaultStyle = new ExcelDefaultStyle();
		//		excelDefaultStyle.createDefaultStyles(workbook, this.excelStyleSet);
	}
	public void writeFormula(String formula)
	{
		createOrGetCell();
		this.cell.setCellType(CellType.NUMERIC);
		this.cell.setCellFormula(formula);
	}
	public void write(double numeric)
	{
		createOrGetCell();

		this.cell.setCellType(CellType.NUMERIC);
		this.cell.setCellValue(numeric);
	}

	public void write(String string) throws IOException
	{
		createOrGetCell();
		//System.out.println("Pos at string==()"+getCellAddress()+"/"+string);
		this.cell.setCellType(CellType.STRING);
		this.cell.setCellValue(string);
	}
	public void write(Date date) throws IOException
	{
		createOrGetCell();
		this.cell.setCellType(CellType.NUMERIC);
		this.cell.setCellValue(date);
	}
	public void setStyle(String key)
	{
		this.cell.setCellStyle(this.excelStyleSet.getStyle(key));
	}

	public void createOrGetCell()
	{
		this.row = this.sheet.getRow(this.currentRowIndex);
		if (this.row == null)
		{
			this.row = this.sheet.createRow(this.currentRowIndex);
		}
		this.cell = this.row.getCell(this.currentColumnIndex);
		if (this.cell == null)
		{
			this.cell = this.row.createCell(this.currentColumnIndex);
		}
	}
	public void writeWorkbook() throws IOException
	{
		this.workbook.write(this.fileOutputStream);
	}
	public void save() throws IOException
	{
		FileOutputStream fileOutputStream = null;
		try
		{
			fileOutputStream = new FileOutputStream(fileNameWithPath);
			this.workbook.write(fileOutputStream);
		}
		catch (Exception e)
		{
			e.printStackTrace();

		}
		finally
		{
			fileOutputStream.close();
		}
	}
	public void createSheet(String sheetName)
	{
		this.sheet = this.workbook.createSheet(sheetName);
	}
	public void createRow(int currentRowIndex)
	{
		this.row = this.sheet.createRow(currentRowIndex);
	}
	public void getSheet(String sheetName)
	{
		this.sheet = this.workbook.getSheet(sheetName);
	}
	public void setMergedRegion(int firstRow, int lastRow, int firstCol, int lastCol)
	{
		this.sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}
	public void setRowHeight(int height)
	{
		this.row = this.sheet.getRow(this.currentRowIndex);
		if (this.row == null)
		{
			this.row = this.sheet.createRow(this.currentRowIndex);
		}
		this.row.setHeight((short)height);
	}
	
	public void addValidationData(DataValidation dataValidation)
	{
		this.sheet.addValidationData(dataValidation);
	}

}
