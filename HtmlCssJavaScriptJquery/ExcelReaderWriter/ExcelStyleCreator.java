package excelreader;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelStyleCreator
{
	public void addExcelStyleSet(Workbook workbook, ExcelStyleSet excelStyleSet);
}
