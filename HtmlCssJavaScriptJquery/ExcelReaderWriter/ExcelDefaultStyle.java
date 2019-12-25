package excelreader;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import com.netx.General;

public class ExcelDefaultStyle implements ExcelStyleCreator
{

	HSSFPalette	palette;
	long		currencyFormatMasterId;
	String curremcySymbol;
	int decimalAllow;

	public int getDecimalAllow()
	{
		return decimalAllow;
	}

	public void setDecimalAllow(int decimalAllow)
	{
		this.decimalAllow = decimalAllow;
	}
	private String getNumberFormat()
	{
		String format = "";
		if(getDecimalAllow()>0)
		{
		if (getCurrencyFormatMasterId() == 0 || getCurrencyFormatMasterId() == 1)
		{
			
			format = General.getFormatUS(getDecimalAllow());
		}
		else
		{
			format = General.getFormatUK( getDecimalAllow());
		}
		}
		return format;
	}
	public static String getCurrencyFormatUS(String USFormat, int maxPrecision)
	{
		if (maxPrecision == 0)
			return USFormat;
		else
		{
			USFormat += "0.";
			for (int i = 0; i < maxPrecision; i++)
			{
				USFormat += "0";
			}
			return USFormat;
		}
	}
	public static String getCurrencyFormatUK(String UKFormat, int maxPrecision)
	{
		if (maxPrecision == 0)
			return UKFormat;
		else
		{
			UKFormat += "0,";
			for (int i = 0; i < maxPrecision; i++)
			{
				UKFormat += "0";
			}
			return UKFormat;
		}
	}
	private String getCurrencyFormat(byte currencyFormatMasterId, String USFormat, String UKFormat, int maxPrecision)
	{
		String format = "";
		if (currencyFormatMasterId == 0 || currencyFormatMasterId == 1)
		{
			format = this.getCurrencyFormatUS(USFormat, maxPrecision);
		}
		else
		{
			format = this.getCurrencyFormatUK(UKFormat, maxPrecision);
		}
		return format;
	}
	private String getCurrencyFormat()
	{
		String format = "";
		if(getCurrencyFormatMasterSymbol()!=null)
		{
		if (getCurrencyFormatMasterId() == 0 || getCurrencyFormatMasterId() == 1)
		{
			
			format = getCurrencyFormatMasterSymbol()+General.getFormatUS(getDecimalAllow());
		}
		else
		{
			format = getCurrencyFormatMasterSymbol()+General.getFormatUK( getDecimalAllow());
		}}
		
		return format;
	}
	
	
	public long getCurrencyFormatMasterId()
	{
		return currencyFormatMasterId;
	}

	public void setCurrencyFormatMasterId(long currencyFormatMasterId)
	{
		this.currencyFormatMasterId = currencyFormatMasterId;
	}
	public String getCurrencyFormatMasterSymbol()
	{
		return curremcySymbol;
	}

	public void setCurrencyFormatMasterSymbol(String curremcySymbol)
	{
		this.curremcySymbol = curremcySymbol;
	}

	public short getHSSFColor(int index, int red, int green, int blue)
	{
		palette.setColorAtIndex((short) index, (byte) red, (byte) green, (byte) blue);
		return palette.getColor(index).getIndex();
	}

	private void createDefaultStyles(Workbook workbook, ExcelStyleSet excelStyleSet)
	{
		//Alka
		String currency_format=this.getCurrencyFormat(); 
		//= this.getCurrencyFormat((byte) 0, decimalAllow);
		String USFormat = "$#,##";
		String UKFormat = "$#.##";

		String kWh = "#,## \"kWh\"";
		String therms_CCF = "#,## \"Therms/CCF\"";
		String perscent_format = "0%";
		String perscent_formatTwoDecimal = "0.00%";

		

		//String currency_format = "";
		String currency_formatTwoDecimal = "";

		String currency_formatFiveDecimal = "";

		String number_format = "";
		
		String currency_formatWithoutDecimalAndDollar = "";

		//				General.getPercentageFormat(4); 
		//				"#.##%";
		String date_format_mmmddyyyy = "mmm dd, yyyy";
		String date_format_mmddyy = "mm/dd/yy";
		String date_format_mmmyyyy = "mmm, yyyy";

		palette = ((HSSFWorkbook) workbook).getCustomPalette();

		CellStyle titleStyle = workbook.createCellStyle();
		/*CellStyle titleStyle1 = workbook.createCellStyle();
		CellStyle headingStyle = workbook.createCellStyle();
		CellStyle cellData = workbook.createCellStyle();*/
		CellStyle cellDataWithCurrencyFormat = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatTwoDecimal = workbook.createCellStyle();

		CellStyle cellDataWithCurrencyFormatTwoDecimalWithRed = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatTwoDecimalWithGreen = workbook.createCellStyle();

		CellStyle cellDataWithCurrencyFormatFiveDecimalWithRed = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatFiveDecimalWithGreen = workbook.createCellStyle();

		CellStyle cellDataWithCurrencyFormatFiveDecimal = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormatWithCommaNoDecimal = workbook.createCellStyle();
		//		CellStyle cellDataWithPercentageFormat = workbook.createCellStyle();
		//CellStyle cellDataWithCurrencyFormat1 = workbook.createCellStyle();
		CellStyle cellDataWithNumberFormat = workbook.createCellStyle();
		CellStyle cellDataWithDateFormat = workbook.createCellStyle();
		CellStyle cellDataWithDateFormat_mmmyyyy = workbook.createCellStyle();

		CellStyle cellDataWithDateFormat2 = workbook.createCellStyle();
		CellStyle cellDataWithDateFormatWithRightAlign = workbook.createCellStyle();
		CellStyle cellDataWithFormatWithLeftAlign = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormat_kWh = workbook.createCellStyle();
		CellStyle cellDataWithNumberFormat_therms_CCF = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormat_perscent_format = workbook.createCellStyle();
		CellStyle cellDataWithNumberFormat_perscent_formatTwoDecimal = workbook.createCellStyle();

		/* New Style Create strike Name */

		CellStyle cellDataWithDateFormat_mmmyyyyWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithDateFormatWithStrikeOut = workbook.createCellStyle();

		CellStyle cellDataWithDateFormatWithRightAlignStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithFormatWithLeftAlignWithStrikeOut = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormatWithCommaNoDecimalStrikeOut = workbook.createCellStyle();

		CellStyle cellDataWithCurrencyFormatWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatTwoDecimalWithStrikeOut = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormatWithCommaNoDecimalWithBold = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatWithBold = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatTwoDecimalWithBold = workbook.createCellStyle();
		CellStyle cellDataWithFormatWithLeftAlignWithBold = workbook.createCellStyle();

		CellStyle cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatWithBoldWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut = workbook.createCellStyle();
		CellStyle cellDataWithCurrencyFormat1 = workbook.createCellStyle();

		Font cellFontStrikeOut = workbook.createFont();
		cellFontStrikeOut.setFontName("ARIAL");
		cellFontStrikeOut.setBold(true);
		cellFontStrikeOut.setFontHeightInPoints((short) 11);
		cellFontStrikeOut.setStrikeout(true);

		Font cellFontStrikeOutWithoutBold = workbook.createFont();
		cellFontStrikeOutWithoutBold.setFontName("ARIAL");
		cellFontStrikeOutWithoutBold.setBold(false);
		cellFontStrikeOutWithoutBold.setFontHeightInPoints((short) 10);
		cellFontStrikeOutWithoutBold.setStrikeout(true);

		Font cellFontWithBold = workbook.createFont();
		cellFontWithBold.setFontName("ARIAL");
		cellFontWithBold.setBold(true);
		cellFontWithBold.setFontHeightInPoints((short) 10);

		Font cellFontWithBoldWithStrikeOut = workbook.createFont();
		cellFontWithBoldWithStrikeOut.setFontName("ARIAL");
		cellFontWithBoldWithStrikeOut.setBold(true);
		cellFontWithBoldWithStrikeOut.setFontHeightInPoints((short) 10);
		cellFontWithBoldWithStrikeOut.setStrikeout(true);

		/* New Style Create strike Name */

		//CellStyle cellDataWithNumberFormat1 = workbook.createCellStyle();
		/*CellStyle cellData1 = workbook.createCellStyle();
		CellStyle totalStyle = workbook.createCellStyle();
		
		CellStyle totalStyle1 = workbook.createCellStyle();
		CellStyle totalStyleWithCurrencyFormat = workbook.createCellStyle();
		*/
		CreationHelper ch = workbook.getCreationHelper();
		DataFormat dataFormat = workbook.createDataFormat();
		//title Style
		Font titleFont = workbook.createFont();
		titleFont.setFontName("ARIAL");
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		titleFont.setFontHeightInPoints((short) 11);

		titleStyle.setFont(titleFont);
		titleStyle.setAlignment(CellStyle.ALIGN_LEFT);
		titleStyle.setWrapText(false);
		titleStyle.setFillForegroundColor(getHSSFColor(HSSFColor.YELLOW.index, 227, 226, 143));
		titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		titleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		titleStyle.setBorderTop(CellStyle.BORDER_THIN);
		titleStyle.setBorderLeft(CellStyle.BORDER_THIN);
		titleStyle.setBorderRight(CellStyle.BORDER_THIN);

		//titleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		//titleStyle.setBorderTop(CellStyle.BORDER_THIN);

		//title style 1
		/*Font titleFont1 = workbook.createFont();
		titleFont1.setFontName("ARIAL");
		//titleFont1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		titleFont1.setFontHeightInPoints((short) 10);
		
		titleStyle1.setFont(titleFont1);
		titleStyle1.setAlignment(CellStyle.ALIGN_LEFT);
		titleStyle1.setWrapText(false);
		titleStyle1.setFillForegroundColor(getHSSFColor(HSSFColor.YELLOW.index, 227, 226, 143));
		titleStyle1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		titleStyle1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		//Heading Style
		Font headingFont = workbook.createFont();
		headingFont.setFontName("ARIAL");
		headingFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headingFont.setFontHeightInPoints((short) 10);
		
		headingStyle.setFont(headingFont);
		headingStyle.setAlignment(CellStyle.ALIGN_LEFT);
		headingStyle.setWrapText(false);
		headingStyle.setFillForegroundColor(getHSSFColor(HSSFColor.LIGHT_YELLOW.index, 249, 249, 227));
		headingStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headingStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		headingStyle.setAlignment(CellStyle.ALIGN_CENTER);
		//headingStyle.setBorderBottom(CellStyle.BORDER_THIN);
		//headingStyle.setBorderTop(CellStyle.BORDER_THIN);
		*/
		//cell data
		Font cellFont = workbook.createFont();
		cellFont.setFontName("ARIAL");
		cellFont.setBold(false);
		cellFont.setFontHeightInPoints((short) 10);

		Font cellFontRed = workbook.createFont();
		cellFontRed.setFontName("ARIAL");
		cellFontRed.setBold(false);
		cellFontRed.setFontHeightInPoints((short) 10);
		cellFontRed.setColor(IndexedColors.RED.getIndex());

		Font cellFontGreen = workbook.createFont();
		cellFontGreen.setFontName("ARIAL");
		cellFontGreen.setBold(false);
		cellFontGreen.setFontHeightInPoints((short) 10);
		cellFontGreen.setColor(IndexedColors.GREEN.getIndex());

		/*cellData.setFont(cellFont);
		cellData.setAlignment(CellStyle.ALIGN_LEFT);
		cellData.setWrapText(false);
		cellData.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellData.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellData.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellData.setAlignment(CellStyleCellStyleCellStyle.ALIGN_CENTER);*/

		// cellDataWithCurrencyFormat
		cellDataWithCurrencyFormat.setFont(cellFont);
		cellDataWithCurrencyFormat.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormat.setWrapText(false);
		cellDataWithCurrencyFormat.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormat.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormat.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormat.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormat.setDataFormat(dataFormat.getFormat(currency_format));
		cellDataWithCurrencyFormat.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormat.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormat.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormat.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatTwoDecimal.setFont(cellFont);
		cellDataWithCurrencyFormatTwoDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimal.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimal.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimal.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimal.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimal.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimal.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimal.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimal.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithNumberFormatWithCommaNoDecimal.setFont(cellFont);
		cellDataWithNumberFormatWithCommaNoDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimal.setWrapText(false);
		cellDataWithNumberFormatWithCommaNoDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormatWithCommaNoDecimal.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormatWithCommaNoDecimal.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormatWithCommaNoDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimal.setDataFormat(dataFormat.getFormat(currency_formatWithoutDecimalAndDollar));
		cellDataWithNumberFormatWithCommaNoDecimal.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimal.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimal.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimal.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatTwoDecimalWithRed.setFont(cellFontRed);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimalWithRed.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimalWithRed.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithRed.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatTwoDecimalWithGreen.setFont(cellFontRed);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithGreen.setBorderRight(CellStyle.BORDER_THIN);

		/**************** five Decimal**************************/

		cellDataWithCurrencyFormatFiveDecimalWithRed.setFont(cellFontRed);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setWrapText(false);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatFiveDecimalWithRed.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setDataFormat(dataFormat.getFormat(currency_formatFiveDecimal));
		cellDataWithCurrencyFormatFiveDecimalWithRed.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithRed.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatFiveDecimalWithGreen.setFont(cellFontGreen);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setWrapText(false);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setDataFormat(dataFormat.getFormat(currency_formatFiveDecimal));
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimalWithGreen.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatFiveDecimal.setFont(cellFont);
		cellDataWithCurrencyFormatFiveDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimal.setWrapText(false);
		cellDataWithCurrencyFormatFiveDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatFiveDecimal.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatFiveDecimal.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatFiveDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatFiveDecimal.setDataFormat(dataFormat.getFormat(currency_formatFiveDecimal));
		cellDataWithCurrencyFormatFiveDecimal.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimal.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimal.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatFiveDecimal.setBorderRight(CellStyle.BORDER_THIN);

		/*//cellDataWithPercentageFormat
		cellDataWithPercentageFormat.setFont(cellFont);
		cellDataWithPercentageFormat.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithPercentageFormat.setWrapText(false);
		cellDataWithPercentageFormat.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithPercentageFormat.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithPercentageFormat.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithPercentageFormat.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithPercentageFormat.setDataFormat(dataFormat.getFormat(this.getCurrencyFormat((byte) 0, "#.##%", "#.##%", 4)));
		cellDataWithPercentageFormat.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithPercentageFormat.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithPercentageFormat.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithPercentageFormat.setBorderRight(CellStyle.BORDER_THIN);*/
		// cellDataWithCurrencyFormat1
		/*cellDataWithCurrencyFormat1.setFont(cellFont);
		cellDataWithCurrencyFormat1.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithCurrencyFormat1.setWrapText(false);
		cellDataWithCurrencyFormat1.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 240, 240, 240));
		cellDataWithCurrencyFormat1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormat1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormat1.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithCurrencyFormat1.setDataFormat(ch.createDataFormat().getFormat(number_format);*/

		//cellDataWithNumberFormat.setDataFormat(df.getFormat(number_format));
		// cellDataWithNumberFormat

		cellDataWithNumberFormat.setFont(cellFont);
		cellDataWithNumberFormat.setWrapText(false);
		cellDataWithNumberFormat.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat.setAlignment(CellStyle.ALIGN_RIGHT);

		cellDataWithNumberFormat_kWh.setFont(cellFont);
		cellDataWithNumberFormat_kWh.setWrapText(false);
		cellDataWithNumberFormat_kWh.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat_kWh.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat_kWh.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat_kWh.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_kWh.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_kWh.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_kWh.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_kWh.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormat_kWh.setDataFormat(dataFormat.getFormat(kWh));

		cellDataWithNumberFormat_therms_CCF.setFont(cellFont);
		cellDataWithNumberFormat_therms_CCF.setWrapText(false);
		cellDataWithNumberFormat_therms_CCF.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat_therms_CCF.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat_therms_CCF.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat_therms_CCF.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_therms_CCF.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_therms_CCF.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_therms_CCF.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_therms_CCF.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormat_therms_CCF.setDataFormat(dataFormat.getFormat(therms_CCF));

		cellDataWithNumberFormat_perscent_format.setFont(cellFont);
		cellDataWithNumberFormat_perscent_format.setWrapText(false);
		cellDataWithNumberFormat_perscent_format.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat_perscent_format.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat_perscent_format.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat_perscent_format.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_format.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_format.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_format.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_format.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormat_perscent_format.setDataFormat(dataFormat.getFormat(perscent_format));

		cellDataWithNumberFormat_perscent_formatTwoDecimal.setFont(cellFont);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setWrapText(false);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormat_perscent_formatTwoDecimal.setDataFormat(dataFormat.getFormat(perscent_formatTwoDecimal));

		cellDataWithFormatWithLeftAlign.setFont(cellFont);
		cellDataWithFormatWithLeftAlign.setWrapText(false);
		cellDataWithFormatWithLeftAlign.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithFormatWithLeftAlign.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithFormatWithLeftAlign.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithFormatWithLeftAlign.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlign.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlign.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlign.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlign.setAlignment(CellStyle.ALIGN_LEFT);

		cellDataWithDateFormat.setFont(cellFont);
		cellDataWithDateFormat.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithDateFormat.setWrapText(false);
		cellDataWithDateFormat.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormat.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormat.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormat.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithDateFormat.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormat.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormat.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormat.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormat.setDataFormat(dataFormat.getFormat(number_format));
		cellDataWithDateFormat.setDataFormat(dataFormat.getFormat(date_format_mmmddyyyy));

		cellDataWithDateFormat_mmmyyyy.setFont(cellFont);
		cellDataWithDateFormat_mmmyyyy.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithDateFormat_mmmyyyy.setWrapText(false);
		cellDataWithDateFormat_mmmyyyy.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormat_mmmyyyy.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormat_mmmyyyy.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormat_mmmyyyy.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithDateFormat_mmmyyyy.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyy.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyy.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyy.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyy.setDataFormat(dataFormat.getFormat(number_format));
		cellDataWithDateFormat_mmmyyyy.setDataFormat(dataFormat.getFormat(date_format_mmmyyyy));

		cellDataWithDateFormat2.setFont(cellFont);
		cellDataWithDateFormat2.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithDateFormat2.setWrapText(false);
		cellDataWithDateFormat2.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormat2.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormat2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormat2.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithDateFormat2.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormat2.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormat2.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormat2.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormat2.setDataFormat(dataFormat.getFormat(number_format));
		cellDataWithDateFormat2.setDataFormat(dataFormat.getFormat(date_format_mmddyy));

		cellDataWithDateFormatWithRightAlign.setFont(cellFont);
		cellDataWithDateFormatWithRightAlign.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithDateFormatWithRightAlign.setWrapText(false);
		cellDataWithDateFormatWithRightAlign.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormatWithRightAlign.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormatWithRightAlign.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormatWithRightAlign.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithDateFormatWithRightAlign.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlign.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlign.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlign.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlign.setDataFormat(dataFormat.getFormat(date_format_mmmddyyyy));

		/*************       New Style StrikeOut  Start ***********************/

		cellDataWithFormatWithLeftAlignWithStrikeOut.setFont(cellFontStrikeOut);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setWrapText(false);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithFormatWithLeftAlignWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithStrikeOut.setAlignment(CellStyle.ALIGN_LEFT);

		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setWrapText(false);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setDataFormat(dataFormat.getFormat(number_format));
		cellDataWithDateFormat_mmmyyyyWithStrikeOut.setDataFormat(dataFormat.getFormat(date_format_mmmyyyy));

		cellDataWithDateFormatWithStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithDateFormatWithStrikeOut.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithDateFormatWithStrikeOut.setWrapText(false);
		cellDataWithDateFormatWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormatWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormatWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormatWithStrikeOut.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithDateFormatWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithStrikeOut.setDataFormat(dataFormat.getFormat(number_format));
		cellDataWithDateFormatWithStrikeOut.setDataFormat(dataFormat.getFormat(date_format_mmddyy));

		cellDataWithDateFormatWithRightAlignStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithDateFormatWithRightAlignStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithDateFormatWithRightAlignStrikeOut.setWrapText(false);
		cellDataWithDateFormatWithRightAlignStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithDateFormatWithRightAlignStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithDateFormatWithRightAlignStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithDateFormatWithRightAlignStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithDateFormatWithRightAlignStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlignStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlignStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlignStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithDateFormatWithRightAlignStrikeOut.setDataFormat(dataFormat.getFormat(date_format_mmmddyyyy));

		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setWrapText(false);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setDataFormat(dataFormat.getFormat(currency_formatWithoutDecimalAndDollar));
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalStrikeOut.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatWithStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithCurrencyFormatWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithStrikeOut.setWrapText(false);
		cellDataWithCurrencyFormatWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithStrikeOut.setDataFormat(dataFormat.getFormat(currency_format));
		cellDataWithCurrencyFormatWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setFont(cellFontStrikeOutWithoutBold);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithFormatWithLeftAlignWithBold.setFont(cellFontWithBold);
		cellDataWithFormatWithLeftAlignWithBold.setWrapText(false);
		cellDataWithFormatWithLeftAlignWithBold.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithFormatWithLeftAlignWithBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithFormatWithLeftAlignWithBold.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithFormatWithLeftAlignWithBold.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBold.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBold.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBold.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBold.setAlignment(CellStyle.ALIGN_LEFT);

		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setFont(cellFontWithBold);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setWrapText(false);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setDataFormat(dataFormat.getFormat(currency_formatWithoutDecimalAndDollar));
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBold.setBorderRight(CellStyle.BORDER_THIN);

		//alka
		cellDataWithCurrencyFormatWithBold.setFont(cellFontWithBold);
		cellDataWithCurrencyFormatWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithBold.setWrapText(false);
		cellDataWithCurrencyFormatWithBold.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatWithBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatWithBold.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithBold.setDataFormat(dataFormat.getFormat(currency_format));
		cellDataWithCurrencyFormatWithBold.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBold.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBold.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBold.setBorderRight(CellStyle.BORDER_THIN);

		
		
		cellDataWithCurrencyFormatTwoDecimalWithBold.setFont(cellFontWithBold);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimalWithBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimalWithBold.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBold.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setFont(cellFontWithBoldWithStrikeOut);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setWrapText(false);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_LEFT);

		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setFont(cellFontWithBoldWithStrikeOut);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setWrapText(false);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setDataFormat(dataFormat.getFormat(currency_formatWithoutDecimalAndDollar));
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setFont(cellFontWithBoldWithStrikeOut);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setWrapText(false);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setDataFormat(dataFormat.getFormat(currency_format));
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatWithBoldWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);

		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setFont(cellFontWithBoldWithStrikeOut);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setWrapText(false);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setAlignment(CellStyle.ALIGN_RIGHT);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setDataFormat(dataFormat.getFormat(currency_formatTwoDecimal));
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setBorderBottom(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setBorderTop(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setBorderLeft(CellStyle.BORDER_THIN);
		cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut.setBorderRight(CellStyle.BORDER_THIN);
		
		
		CellStyle headingStyle_template = workbook.createCellStyle();
		headingStyle_template.setFont(titleFont);
		headingStyle_template.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		headingStyle_template.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headingStyle_template.setWrapText(false);
		headingStyle_template.setFillBackgroundColor(HSSFColor.BLUE.index);
		headingStyle_template.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		/*************       New Style StrikeOut  end ***********************/

		//dataH2WithDateStyle.setDataFormat(ch.createDataFormat().getFormat("mmm dd, yyyy h:mm AM/PM"));

		// cellDataWithNumberFormat1
		/*cellDataWithNumberFormat1.setFont(cellFont);
		cellDataWithNumberFormat1.setAlignment(CellStyle.ALIGN_LEFT);
		cellDataWithNumberFormat1.setWrapText(false);
		cellDataWithNumberFormat1.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 240, 240, 240));
		cellDataWithNumberFormat1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellDataWithNumberFormat1.setAlignment(CellStyle.ALIGN_CENTER);
		cellDataWithNumberFormat1.setDataFormat(df.getFormat(number_format));*/
		/*
				//cell data1
				Font cellFont1 = workbook.createFont();
				cellFont1.setFontName("ARIAL");
				cellFont1.setBold(false);
				cellFont1.setFontHeightInPoints((short) 10);
		
				cellData1.setFont(cellFont);
				cellData1.setAlignment(CellStyle.ALIGN_LEFT);
				cellData1.setWrapText(false);
				cellData1.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 240, 240, 240));
				cellData1.setFillPattern(CellStyle.SOLID_FOREGROUND);
				cellData1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				cellData1.setAlignment(CellStyle.ALIGN_CENTER);
		
				// style for total auction heading 
				Font totalFont = workbook.createFont();
				totalFont.setFontName("ARIAL");
				totalFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				totalFont.setFontHeightInPoints((short) 9);
		
				totalStyle.setFont(totalFont);
				totalStyle.setAlignment(CellStyle.ALIGN_LEFT);
				totalStyle.setWrapText(false);
				totalStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 240, 240, 240));
				totalStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				totalStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				totalStyle.setAlignment(CellStyle.ALIGN_CENTER);
		
				//style for total auction details
				Font totalFont1 = workbook.createFont();
				totalFont1.setFontName("ARIAL");
				totalFont1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				totalFont1.setFontHeightInPoints((short) 10);
		
				totalStyle1.setFont(totalFont1);
				totalStyle1.setAlignment(CellStyle.ALIGN_LEFT);
				totalStyle1.setWrapText(false);
				totalStyle1.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_40_PERCENT.index, 240, 240, 240));
				totalStyle1.setFillPattern(CellStyle.SOLID_FOREGROUND);
				totalStyle1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				totalStyle1.setAlignment(CellStyle.ALIGN_CENTER);
		
				// totalStyleWithCurrencyFormat
				totalStyleWithCurrencyFormat.setFont(totalFont1);
				totalStyleWithCurrencyFormat.setAlignment(CellStyle.ALIGN_LEFT);
				totalStyleWithCurrencyFormat.setWrapText(false);
				totalStyleWithCurrencyFormat.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_40_PERCENT.index, 240, 240, 240));
				totalStyleWithCurrencyFormat.setFillPattern(CellStyle.SOLID_FOREGROUND);
				totalStyleWithCurrencyFormat.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				totalStyleWithCurrencyFormat.setAlignment(CellStyle.ALIGN_CENTER);
				totalStyleWithCurrencyFormat.setDataFormat(df.getFormat(currency_format));
		*/
		// set hashmap for style
		excelStyleSet.setStyle("titleStyle", titleStyle);
		excelStyleSet.setStyle("headingStyle_template",headingStyle_template);
		/*excelStyleSet.setStyle("titleStyle1", titleStyle1);
		excelStyleSet.setStyle("headingStyle", headingStyle);
		excelStyleSet.setStyle("cellData", cellData);*/
		excelStyleSet.setStyle("cellDataWithCurrencyFormat", cellDataWithCurrencyFormat);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimal", cellDataWithCurrencyFormatTwoDecimal);

		excelStyleSet.setStyle("cellDataWithCurrencyFormatFiveDecimalWithGreen", cellDataWithCurrencyFormatFiveDecimalWithGreen);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatFiveDecimalWithRed", cellDataWithCurrencyFormatFiveDecimalWithRed);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimalWithRed", cellDataWithCurrencyFormatTwoDecimalWithRed);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimalWithGreen", cellDataWithCurrencyFormatTwoDecimalWithGreen);

		excelStyleSet.setStyle("cellDataWithCurrencyFormatFiveDecimal", cellDataWithCurrencyFormatFiveDecimal);
		excelStyleSet.setStyle("cellDataWithNumberFormatWithCommaNoDecimal", cellDataWithNumberFormatWithCommaNoDecimal);

		//		excelStyleSet.setStyle("cellDataWithPercentageFormat", cellDataWithPercentageFormat);
		//excelStyleSet.setStyle("cellDataWithCurrencyFormat1", cellDataWithCurrencyFormat1);
		excelStyleSet.setStyle("cellDataWithNumberFormat", cellDataWithNumberFormat);
		excelStyleSet.setStyle("cellDataWithDateFormat", cellDataWithDateFormat);
		excelStyleSet.setStyle("cellDataWithDateFormatWithRightAlign", cellDataWithDateFormatWithRightAlign);
		excelStyleSet.setStyle("cellDataWithFormatWithLeftAlign", cellDataWithFormatWithLeftAlign);
		excelStyleSet.setStyle("cellDataWithDateFormat2", cellDataWithDateFormat2);

		excelStyleSet.setStyle("cellDataWithDateFormat_mmmyyyy", cellDataWithDateFormat_mmmyyyy);

		excelStyleSet.setStyle("cellDataWithNumberFormat_kWh", cellDataWithNumberFormat_kWh);
		excelStyleSet.setStyle("cellDataWithNumberFormat_therms_CCF", cellDataWithNumberFormat_therms_CCF);
		excelStyleSet.setStyle("cellDataWithNumberFormat_perscent_format", cellDataWithNumberFormat_perscent_format);
		excelStyleSet.setStyle("cellDataWithNumberFormat_perscent_formatTwoDecimal", cellDataWithNumberFormat_perscent_formatTwoDecimal);

		/******************************** new Style for StrikeOut Start **************************************/

		excelStyleSet.setStyle("cellDataWithFormatWithLeftAlignWithStrikeOut", cellDataWithFormatWithLeftAlignWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithDateFormat_mmmyyyyWithStrikeOut", cellDataWithDateFormat_mmmyyyyWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithDateFormatWithStrikeOut", cellDataWithDateFormatWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithDateFormatWithRightAlignStrikeOut", cellDataWithDateFormatWithRightAlignStrikeOut);
		excelStyleSet.setStyle("cellDataWithNumberFormatWithCommaNoDecimalStrikeOut", cellDataWithNumberFormatWithCommaNoDecimalStrikeOut);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatWithStrikeOut", cellDataWithCurrencyFormatWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimalWithStrikeOut", cellDataWithCurrencyFormatTwoDecimalWithStrikeOut);

		excelStyleSet.setStyle("cellDataWithNumberFormatWithCommaNoDecimalWithBold", cellDataWithNumberFormatWithCommaNoDecimalWithBold);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatWithBold", cellDataWithCurrencyFormatWithBold);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimalWithBold", cellDataWithCurrencyFormatTwoDecimalWithBold);
		excelStyleSet.setStyle("cellDataWithFormatWithLeftAlignWithBold", cellDataWithFormatWithLeftAlignWithBold);

		excelStyleSet.setStyle("cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut", cellDataWithNumberFormatWithCommaNoDecimalWithBoldWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatWithBoldWithStrikeOut", cellDataWithCurrencyFormatWithBoldWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut", cellDataWithCurrencyFormatTwoDecimalWithBoldWithStrikeOut);
		excelStyleSet.setStyle("cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut", cellDataWithFormatWithLeftAlignWithBoldWithStrikeOut);

		/******************************** new Style for StrikeOut End **************************************/

		//excelStyleSet.setStyle("cellDataWithNumberFormat1", cellDataWithNumberFormat1);
		/*excelStyleSet.setStyle("cellData1", cellData1);
		excelStyleSet.setStyle("totalStyle", totalStyle);
		excelStyleSet.setStyle("totalStyle1", totalStyle1);
		excelStyleSet.setStyle("totalStyleWithCurrencyFormat", totalStyleWithCurrencyFormat);
		*/ }

	@Override
	public void addExcelStyleSet(Workbook workbook, ExcelStyleSet excelStyleSet)
	{
		///////////////////
		palette = ((HSSFWorkbook) workbook).getCustomPalette();
		DataFormat dataFormat = workbook.createDataFormat();
		CreationHelper ch = workbook.getCreationHelper();
		String date_format_MMM_yy = "MMM-yy";
		String currency_format ="";
		String percentage_format ="";
		if (this.getCurrencyFormat() != null)
		{
			currency_format = this.getCurrencyFormat();

		}
		String number_format = "";
		if (this.getNumberFormat() != null)
		{
			number_format = this.getNumberFormat();
		}

		String CurrencySymbol = "";
		if (this.getCurrencyFormatMasterSymbol() != null)
		{
			CurrencySymbol = this.getCurrencyFormatMasterSymbol();
		}
		String number_format3 = "#,##0";
		String number_format1 = CurrencySymbol + "#,##0.000";
		String number_format2 = CurrencySymbol + number_format3;
		int decimalAllowed = this.getDecimalAllow();
		percentage_format=General.getFormatUSForPercentageForNaturalGas(decimalAllowed);
		///////////////////
		// For Formating of Cell Style
				//int decimalAllowed = 5;
				String format="";
				String formatPercentage="";
				String subFormat="";
				if(this.getDecimalAllow()>0)
				{
					format="##,###,###,###,##0.";
					for(int i=0;i<this.getDecimalAllow();i++)
					{
						subFormat+="0";
					}
					format+=subFormat;
				}
				else
				{
					format="##,###,###,###,##0.00";
				}
				formatPercentage="##,###,###,###,##0.00%";
				String perscent_formatTwoDecimal = format;
		///////////////////////	
		//////////////////
		//TitleStyle///////////////////////////////////////
		Font titleFont = workbook.createFont();
		titleFont.setFontName("ARIAL");
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		titleFont.setFontHeightInPoints((short) 12);
		titleFont.setColor(HSSFColor.WHITE.index);
		
		CellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFont(titleFont);
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		titleStyle.setWrapText(false);
		titleStyle.setFillBackgroundColor(HSSFColor.BLACK.index);
		titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
		
		excelStyleSet.setStyle("titleStyle", titleStyle);	
		
		//////////////////////////////////////
		
		
		//////////////////
		//HeadingStyle///////////////////////////////////////
		Font headingFont = workbook.createFont();
		headingFont.setFontName("ARIAL");
		headingFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headingFont.setFontHeightInPoints((short) 12);
		headingFont.setColor(HSSFColor.BLACK.index);
		
		CellStyle headingStyle = workbook.createCellStyle();
		headingStyle = workbook.createCellStyle();
		headingStyle.setFont(headingFont);
		headingStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headingStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		headingStyle.setWrapText(true);
		headingStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		headingStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headingStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headingStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headingStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headingStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headingStyle", headingStyle);		
		
		
		//////////////////////////////////////
		
		
		//////////////////
		//CellDataStyle///////////////////////////////////////
		
		
		Font cellDataFont = workbook.createFont();
		cellDataFont.setFontName("Areal");
		cellDataFont.setFontHeightInPoints((short) 10);
		cellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		cellDataFont.setColor(HSSFColor.BLACK.index);
		
		Font cellDataFont1 = workbook.createFont();
		cellDataFont1.setFontName("Areal");
		cellDataFont1.setFontHeightInPoints((short) 10);
		cellDataFont1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		cellDataFont1.setColor(HSSFColor.WHITE.index);
		
		
		CellStyle cellDataStyle = workbook.createCellStyle();
		cellDataStyle = workbook.createCellStyle();
		cellDataStyle.setFont(cellDataFont);
		cellDataStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		cellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		cellDataStyle.setWrapText(true);
		cellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_40_PERCENT.index, 217, 217, 217));
		cellDataStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("cellDataStyle", cellDataStyle);	
		//////////////////////////////////////
		///template work
		Font headingCellDataFont = workbook.createFont();
		headingCellDataFont.setFontName("Areal");
		headingCellDataFont.setFontHeightInPoints((short) 11);
		headingCellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headingCellDataFont.setColor(HSSFColor.BLACK.index);
		
		
		CellStyle headingcellDataStyle = workbook.createCellStyle();
		headingcellDataStyle = workbook.createCellStyle();
		headingcellDataStyle.setFont(headingCellDataFont);
		headingcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headingcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headingcellDataStyle.setWrapText(true);
		headingcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		headingcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		/*headingcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_DASH_DOT);*/
		excelStyleSet.setStyle("headingcellDataStyle", headingcellDataStyle);
		
		Font mainHeadingCellDataFont = workbook.createFont();
		mainHeadingCellDataFont.setFontName("Areal");
		mainHeadingCellDataFont.setFontHeightInPoints((short) 16);
		mainHeadingCellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		mainHeadingCellDataFont.setColor(HSSFColor.BLACK.index);
		
		CellStyle mainHeadingcellDataStyle = workbook.createCellStyle();
		mainHeadingcellDataStyle = workbook.createCellStyle();
		mainHeadingcellDataStyle.setFont(mainHeadingCellDataFont);
		mainHeadingcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		mainHeadingcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		mainHeadingcellDataStyle.setWrapText(true);
		mainHeadingcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		mainHeadingcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		/*headingcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_DASH_DOT);
		headingcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_DASH_DOT);*/
		excelStyleSet.setStyle("mainHeadingcellDataStyle", mainHeadingcellDataStyle);	
		
		CellStyle sectionHeadingcellDataStyle = workbook.createCellStyle();
		sectionHeadingcellDataStyle = workbook.createCellStyle();
		sectionHeadingcellDataStyle.setFont(headingCellDataFont);
		sectionHeadingcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		sectionHeadingcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		sectionHeadingcellDataStyle.setWrapText(true);
		sectionHeadingcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		sectionHeadingcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		sectionHeadingcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		sectionHeadingcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		sectionHeadingcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		sectionHeadingcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("sectionHeadingcellDataStyle", sectionHeadingcellDataStyle);	
		
		
		CellStyle totalHeadingcellDataStyle = workbook.createCellStyle();
		totalHeadingcellDataStyle = workbook.createCellStyle();
		totalHeadingcellDataStyle.setFont(headingCellDataFont);
		totalHeadingcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		totalHeadingcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		totalHeadingcellDataStyle.setWrapText(true);
		totalHeadingcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		totalHeadingcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		totalHeadingcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		totalHeadingcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		totalHeadingcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		totalHeadingcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("totalHeadingcellDataStyle", totalHeadingcellDataStyle);	
		
		
		Font templateCellDataFont = workbook.createFont();
		templateCellDataFont.setFontName("Areal");
		templateCellDataFont.setFontHeightInPoints((short) 10);
		templateCellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		templateCellDataFont.setColor(HSSFColor.BLACK.index);
		
		
		CellStyle templatecellDataStyleOdd = workbook.createCellStyle();
		templatecellDataStyleOdd = workbook.createCellStyle();
		templatecellDataStyleOdd.setFont(templateCellDataFont);
		templatecellDataStyleOdd.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		templatecellDataStyleOdd.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templatecellDataStyleOdd.setWrapText(true);
		templatecellDataStyleOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		templatecellDataStyleOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templatecellDataStyleOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templatecellDataStyleOdd", templatecellDataStyleOdd);	
	
		
		CellStyle templatecellDataStyleEven = workbook.createCellStyle();
		templatecellDataStyleEven = workbook.createCellStyle();
		templatecellDataStyleEven.setFont(templateCellDataFont);
		templatecellDataStyleEven.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		templatecellDataStyleEven.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templatecellDataStyleEven.setWrapText(true);
		templatecellDataStyleEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		templatecellDataStyleEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templatecellDataStyleEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templatecellDataStyleEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templatecellDataStyleEven", templatecellDataStyleEven);	
		
		
		CellStyle templateRightcellDataStyleOdd = workbook.createCellStyle();
		templateRightcellDataStyleOdd = workbook.createCellStyle();
		templateRightcellDataStyleOdd.setFont(templateCellDataFont);
		templateRightcellDataStyleOdd.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		templateRightcellDataStyleOdd.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateRightcellDataStyleOdd.setWrapText(true);
		templateRightcellDataStyleOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		templateRightcellDataStyleOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateRightcellDataStyleOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateRightcellDataStyleOdd", templateRightcellDataStyleOdd);	
	
		
		CellStyle templateRightcellDataStyleEven = workbook.createCellStyle();
		templateRightcellDataStyleEven = workbook.createCellStyle();
		templateRightcellDataStyleEven.setFont(templateCellDataFont);
		templateRightcellDataStyleEven.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		templateRightcellDataStyleEven.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateRightcellDataStyleEven.setWrapText(true);
		templateRightcellDataStyleEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		templateRightcellDataStyleEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateRightcellDataStyleEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateRightcellDataStyleEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateRightcellDataStyleEven", templateRightcellDataStyleEven);	
		
		
		
		Font templateBoldHeadingCellDataFont = workbook.createFont();
		templateBoldHeadingCellDataFont.setFontName("Areal");
		templateBoldHeadingCellDataFont.setFontHeightInPoints((short) 10);
		templateBoldHeadingCellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		templateBoldHeadingCellDataFont.setColor(HSSFColor.BLACK.index);
		
		
		CellStyle templateBoldcellDataStyleOdd = workbook.createCellStyle();
		templateBoldcellDataStyleOdd = workbook.createCellStyle();
		templateBoldcellDataStyleOdd.setFont(templateBoldHeadingCellDataFont);
		templateBoldcellDataStyleOdd.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBoldcellDataStyleOdd.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBoldcellDataStyleOdd.setWrapText(true);
		templateBoldcellDataStyleOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		templateBoldcellDataStyleOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateBoldcellDataStyleOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateBoldcellDataStyleOdd", templateBoldcellDataStyleOdd);	
				
		
		CellStyle templateBoldcellDataStyleEven = workbook.createCellStyle();
		templateBoldcellDataStyleEven = workbook.createCellStyle();
		templateBoldcellDataStyleEven.setFont(templateBoldHeadingCellDataFont);
		templateBoldcellDataStyleEven.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBoldcellDataStyleEven.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBoldcellDataStyleEven.setWrapText(true);
		templateBoldcellDataStyleEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		templateBoldcellDataStyleEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateBoldcellDataStyleEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateBoldcellDataStyleEven", templateBoldcellDataStyleEven);	
		
		
		Font templateNormalHeadingCellDataFont = workbook.createFont();
		templateNormalHeadingCellDataFont.setFontName("Areal");
		templateNormalHeadingCellDataFont.setFontHeightInPoints((short) 11);
		templateNormalHeadingCellDataFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		templateNormalHeadingCellDataFont.setColor(HSSFColor.BLACK.index);
	
		
		Font templateNormalHeadingCellDataFont_benchmark = workbook.createFont();
		templateNormalHeadingCellDataFont_benchmark.setFontName("Areal");
		templateNormalHeadingCellDataFont_benchmark.setFontHeightInPoints((short) 10);
		templateNormalHeadingCellDataFont_benchmark.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		templateNormalHeadingCellDataFont_benchmark.setColor(HSSFColor.BLACK.index);
		
		CellStyle templateNormalcellDataStyle = workbook.createCellStyle();
		templateNormalcellDataStyle = workbook.createCellStyle();
		templateNormalcellDataStyle.setFont(templateNormalHeadingCellDataFont);
		templateNormalcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		templateNormalcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateNormalcellDataStyle.setWrapText(true);
		templateNormalcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 244, 244, 244));
		templateNormalcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		templateNormalcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateNormalcellDataStyle", templateNormalcellDataStyle);	
		
		
		CellStyle templateBidderNamecellDataStyleOdd = workbook.createCellStyle();
		templateBidderNamecellDataStyleOdd = workbook.createCellStyle();
		templateBidderNamecellDataStyleOdd.setFont(headingCellDataFont);
		templateBidderNamecellDataStyleOdd.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBidderNamecellDataStyleOdd.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBidderNamecellDataStyleOdd.setWrapText(true);
		templateBidderNamecellDataStyleOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		templateBidderNamecellDataStyleOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateBidderNamecellDataStyleOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBidderNamecellDataStyleOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBidderNamecellDataStyleOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBidderNamecellDataStyleOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateBidderNamecellDataStyleOdd", templateBidderNamecellDataStyleOdd);	
				
		//hiren
		// for bold and right align
		Font templateNormalCellDataFontBold = workbook.createFont();
		templateNormalCellDataFontBold.setFontName("Areal");
		templateNormalCellDataFontBold.setFontHeightInPoints((short) 11);
		templateNormalCellDataFontBold.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		templateNormalCellDataFontBold.setBold(true);
		templateNormalCellDataFontBold.setColor(HSSFColor.BLACK.index);
	
		CellStyle templateNormalcellDataStyleBold = workbook.createCellStyle();
		templateNormalcellDataStyleBold = workbook.createCellStyle();
		templateNormalcellDataStyleBold.setFont(templateNormalCellDataFontBold);
		templateNormalcellDataStyleBold.setAlignment(CellStyle.ALIGN_RIGHT);
		templateNormalcellDataStyleBold.setVerticalAlignment(CellStyle.ALIGN_RIGHT);
		templateNormalcellDataStyleBold.setWrapText(true);
		templateNormalcellDataStyleBold.setFillForegroundColor(getHSSFColor(HSSFColor.BLACK.index, 244, 244, 244));
		templateNormalcellDataStyleBold.setFillPattern(HSSFCellStyle.BORDER_NONE);
		templateNormalcellDataStyleBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyleBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyleBold.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateNormalcellDataStyleBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateNormalcellDataStyleBold", templateNormalcellDataStyleBold);	
		
		
		//  for bold and upto 2 decimal value
		CellStyle NumberStyleWithCurrencyBold  = workbook.createCellStyle();
		Font  columnFont3= workbook.createFont();
		columnFont3.setFontName("Areal");
		columnFont3.setFontHeightInPoints((short) 10);
		columnFont3.setBold(true);
		columnFont3.setColor(HSSFColor.BLACK.index);
		
		CellStyle NumberStyleWithCurrencyOddWithTwoDecimalBold  = workbook.createCellStyle();
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setFont(columnFont3);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setWrapText(false);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0.00"));
		NumberStyleWithCurrencyOddWithTwoDecimalBold .setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyOddWithTwoDecimalBold", NumberStyleWithCurrencyOddWithTwoDecimalBold );	
		
		CellStyle NumberStyleWithCurrencyEvenWithTwoDecimalBold  = workbook.createCellStyle();
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setFont(columnFont3);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setWrapText(false);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0.00"));
		NumberStyleWithCurrencyEvenWithTwoDecimalBold.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyEvenWithTwoDecimalBold", NumberStyleWithCurrencyEvenWithTwoDecimalBold);	
		//end
		CellStyle templateBidderBoldcellDataStyleEven = workbook.createCellStyle();
		templateBidderBoldcellDataStyleEven = workbook.createCellStyle();
		templateBidderBoldcellDataStyleEven.setFont(headingCellDataFont);
		templateBidderBoldcellDataStyleEven.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBidderBoldcellDataStyleEven.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBidderBoldcellDataStyleEven.setWrapText(true);
		templateBidderBoldcellDataStyleEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		templateBidderBoldcellDataStyleEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateBidderBoldcellDataStyleEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBidderBoldcellDataStyleEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBidderBoldcellDataStyleEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBidderBoldcellDataStyleEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateBidderBoldcellDataStyleEven", templateBidderBoldcellDataStyleEven);	
		
		CellStyle NumberStyleNoCurrencyOdd  = workbook.createCellStyle();
		Font  columnFont= workbook.createFont();
		columnFont.setFontName("Areal");
		columnFont.setFontHeightInPoints((short) 10);
		columnFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		columnFont.setColor(HSSFColor.BLACK.index);
		
		NumberStyleNoCurrencyOdd.setFont(columnFont);
		NumberStyleNoCurrencyOdd.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleNoCurrencyOdd.setWrapText(false);
		NumberStyleNoCurrencyOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleNoCurrencyOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleNoCurrencyOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd.setDataFormat(ch.createDataFormat().getFormat("##,###,###,###,##0"));
		NumberStyleNoCurrencyOdd.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleNoCurrencyOdd", NumberStyleNoCurrencyOdd);	
		
		CellStyle NumberStyleNoCurrencyEven  = workbook.createCellStyle();
		NumberStyleNoCurrencyEven.setFont(columnFont);
		NumberStyleNoCurrencyEven.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleNoCurrencyEven.setWrapText(false);
		NumberStyleNoCurrencyEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleNoCurrencyEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleNoCurrencyEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven.setDataFormat(ch.createDataFormat().getFormat("##,###,###,###,##0"));
		NumberStyleNoCurrencyEven.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleNoCurrencyEven", NumberStyleNoCurrencyEven);	
		
		
		CellStyle NumberStyleWithCurrencyOdd  = workbook.createCellStyle();
		Font  columnFont2= workbook.createFont();
		columnFont2.setFontName("Areal");
		columnFont2.setFontHeightInPoints((short) 10);
		columnFont2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		columnFont2.setColor(HSSFColor.BLACK.index);
		
		NumberStyleWithCurrencyOdd.setFont(columnFont2);
		NumberStyleWithCurrencyOdd.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyOdd.setWrapText(false);
		NumberStyleWithCurrencyOdd.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleWithCurrencyOdd.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOdd.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		NumberStyleWithCurrencyOdd.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyOdd", NumberStyleWithCurrencyOdd);	
		
		CellStyle NumberStyleWithCurrencyEven  = workbook.createCellStyle();
		NumberStyleWithCurrencyEven.setFont(columnFont2);
		NumberStyleWithCurrencyEven.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyEven.setWrapText(false);
		NumberStyleWithCurrencyEven.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleWithCurrencyEven.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyEven.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEven.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEven.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEven.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEven.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		NumberStyleWithCurrencyEven.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyEven", NumberStyleWithCurrencyEven);	
		
		CellStyle NumberStyleWithCurrencyOddWithTwoDecimal  = workbook.createCellStyle();
		
		NumberStyleWithCurrencyOddWithTwoDecimal.setFont(columnFont2);
		NumberStyleWithCurrencyOddWithTwoDecimal.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyOddWithTwoDecimal.setWrapText(false);
		NumberStyleWithCurrencyOddWithTwoDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleWithCurrencyOddWithTwoDecimal.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyOddWithTwoDecimal.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimal.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimal.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimal.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithTwoDecimal.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0.00"));
		NumberStyleWithCurrencyOddWithTwoDecimal.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyOddWithTwoDecimal", NumberStyleWithCurrencyOddWithTwoDecimal);	
		
		CellStyle NumberStyleWithCurrencyEvenWithTwoDecimal  = workbook.createCellStyle();
		
		NumberStyleWithCurrencyEvenWithTwoDecimal.setFont(columnFont2);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyEvenWithTwoDecimal.setWrapText(false);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleWithCurrencyEvenWithTwoDecimal.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithTwoDecimal.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0.00"));
		NumberStyleWithCurrencyEvenWithTwoDecimal.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyEvenWithTwoDecimal", NumberStyleWithCurrencyEvenWithTwoDecimal);	
		
		CellStyle templateBlackRow  = workbook.createCellStyle();
		templateBlackRow.setFont(columnFont2);
		templateBlackRow.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		templateBlackRow.setWrapText(true);
		templateBlackRow.setFillForegroundColor(getHSSFColor(HSSFColor.BLACK.index, 0, 0, 0));
		templateBlackRow.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		templateBlackRow.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBlackRow.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBlackRow.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBlackRow.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("templateBlackRow", templateBlackRow);	
		
		CellStyle templateBenchmarkcellDataStyle = workbook.createCellStyle();
		templateBenchmarkcellDataStyle = workbook.createCellStyle();
		templateBenchmarkcellDataStyle.setFont(templateNormalHeadingCellDataFont_benchmark);
		templateBenchmarkcellDataStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		templateBenchmarkcellDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBenchmarkcellDataStyle.setWrapText(false);
		templateBenchmarkcellDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 244, 244, 244));
		templateBenchmarkcellDataStyle.setFillPattern(HSSFCellStyle.BORDER_NONE);
		templateBenchmarkcellDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBenchmarkcellDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBenchmarkcellDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBenchmarkcellDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		templateBenchmarkcellDataStyle.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		//templateBenchmarkcellDataStyle.setDataFormat(dataFormat.getFormat(this.getCurrencyFormat((byte) 1, "#.##", "#.##", 0)));
		//templateBenchmarkcellDataStyle.setDataFormat(dataFormat.getFormat(currency_format));
		templateBenchmarkcellDataStyle.setWrapText(true);
		excelStyleSet.setStyle("templateBenchmarkcellDataStyle", templateBenchmarkcellDataStyle);	
				
		
		
		CellStyle templateBoldcellDataStyleOddWithOutBorder = workbook.createCellStyle();
		templateBoldcellDataStyleOddWithOutBorder = workbook.createCellStyle();
		templateBoldcellDataStyleOddWithOutBorder.setFont(templateBoldHeadingCellDataFont);
		templateBoldcellDataStyleOddWithOutBorder.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBoldcellDataStyleOddWithOutBorder.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBoldcellDataStyleOddWithOutBorder.setWrapText(true);
		templateBoldcellDataStyleOddWithOutBorder.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		templateBoldcellDataStyleOddWithOutBorder.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		/*templateBoldcellDataStyleOdd.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleOdd.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		excelStyleSet.setStyle("templateBoldcellDataStyleOddWithOutBorder", templateBoldcellDataStyleOddWithOutBorder);	
				
		
		CellStyle templateBoldcellDataStyleEvenWithOutBorder = workbook.createCellStyle();
		templateBoldcellDataStyleEvenWithOutBorder = workbook.createCellStyle();
		templateBoldcellDataStyleEvenWithOutBorder.setFont(templateBoldHeadingCellDataFont);
		templateBoldcellDataStyleEvenWithOutBorder.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		templateBoldcellDataStyleEvenWithOutBorder.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		templateBoldcellDataStyleEvenWithOutBorder.setWrapText(true);
		templateBoldcellDataStyleEvenWithOutBorder.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		templateBoldcellDataStyleEvenWithOutBorder.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		/*templateBoldcellDataStyleEvenWithOutBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEvenWithOutBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEvenWithOutBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
		templateBoldcellDataStyleEvenWithOutBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		excelStyleSet.setStyle("templateBoldcellDataStyleEvenWithOutBorder", templateBoldcellDataStyleEvenWithOutBorder);	
		
		
		CellStyle percentageOddDataStyle = workbook.createCellStyle();
		percentageOddDataStyle = workbook.createCellStyle();
		percentageOddDataStyle.setFont(templateNormalHeadingCellDataFont);
		percentageOddDataStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		percentageOddDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		percentageOddDataStyle.setWrapText(false);
		percentageOddDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		percentageOddDataStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		/*percentageOddDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		percentageOddDataStyle.setDataFormat(dataFormat.getFormat(percentage_format));
	//	percentageOddDataStyle.setWrapText(true);
		excelStyleSet.setStyle("percentageOddDataStyle", percentageOddDataStyle);
		
		
		CellStyle percentageEvenDataStyle = workbook.createCellStyle();
		
		percentageEvenDataStyle.setFont(templateNormalHeadingCellDataFont);
		percentageEvenDataStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		percentageEvenDataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		percentageEvenDataStyle.setWrapText(false);
		percentageEvenDataStyle.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		percentageEvenDataStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		/*percentageEvenDataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		percentageEvenDataStyle.setDataFormat(dataFormat.getFormat(percentage_format));
	//	percentageEvenDataStyle.setWrapText(true);
		excelStyleSet.setStyle("percentageEvenDataStyle", percentageEvenDataStyle);
		
		CellStyle NumberStyleWithCurrencyOddWithoutBorder  = workbook.createCellStyle();
		
		
		NumberStyleWithCurrencyOddWithoutBorder.setFont(columnFont2);
		NumberStyleWithCurrencyOddWithoutBorder.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyOddWithoutBorder.setWrapText(false);
		NumberStyleWithCurrencyOddWithoutBorder.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleWithCurrencyOddWithoutBorder.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
/*		NumberStyleWithCurrencyOddWithoutBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyOddWithoutBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);*/
		NumberStyleWithCurrencyOddWithoutBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
		/*NumberStyleWithCurrencyOddWithoutBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		NumberStyleWithCurrencyOddWithoutBorder.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		NumberStyleWithCurrencyOddWithoutBorder.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyOddWithoutBorder", NumberStyleWithCurrencyOddWithoutBorder);	
		
		CellStyle NumberStyleWithCurrencyEvenWithoutBorder  = workbook.createCellStyle();
		NumberStyleWithCurrencyEvenWithoutBorder.setFont(columnFont2);
		NumberStyleWithCurrencyEvenWithoutBorder.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleWithCurrencyEvenWithoutBorder.setWrapText(false);
		NumberStyleWithCurrencyEvenWithoutBorder.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleWithCurrencyEvenWithoutBorder.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	/*	NumberStyleWithCurrencyEvenWithoutBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleWithCurrencyEvenWithoutBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);*/
		NumberStyleWithCurrencyEvenWithoutBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
		/*NumberStyleWithCurrencyEvenWithoutBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		NumberStyleWithCurrencyEvenWithoutBorder.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		NumberStyleWithCurrencyEvenWithoutBorder.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleWithCurrencyEvenWithoutBorder", NumberStyleWithCurrencyEvenWithoutBorder);	
		
		CellStyle HeadingcellDataStyleWithoutBorder = workbook.createCellStyle();
		HeadingcellDataStyleWithoutBorder = workbook.createCellStyle();
		HeadingcellDataStyleWithoutBorder.setFont(headingCellDataFont);
		HeadingcellDataStyleWithoutBorder.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		HeadingcellDataStyleWithoutBorder.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		HeadingcellDataStyleWithoutBorder.setWrapText(true);
		HeadingcellDataStyleWithoutBorder.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		HeadingcellDataStyleWithoutBorder.setFillPattern(HSSFCellStyle.BORDER_NONE);
		/*HeadingcellDataStyleWithoutBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);*/
		HeadingcellDataStyleWithoutBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		/*HeadingcellDataStyleWithoutBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
		HeadingcellDataStyleWithoutBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		excelStyleSet.setStyle("HeadingcellDataStyleWithoutBorder", HeadingcellDataStyleWithoutBorder);	
		
		
		CellStyle SectioncellDataStyleWithleftBorder = workbook.createCellStyle();
		SectioncellDataStyleWithleftBorder = workbook.createCellStyle();
		SectioncellDataStyleWithleftBorder.setFont(headingCellDataFont);
		SectioncellDataStyleWithleftBorder.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		SectioncellDataStyleWithleftBorder.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		SectioncellDataStyleWithleftBorder.setWrapText(true);
		SectioncellDataStyleWithleftBorder.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		SectioncellDataStyleWithleftBorder.setFillPattern(HSSFCellStyle.BORDER_NONE);
		/*SectioncellDataStyleWithleftBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);*/
		SectioncellDataStyleWithleftBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		/*SectioncellDataStyleWithleftBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
		SectioncellDataStyleWithleftBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);*/
		excelStyleSet.setStyle("SectioncellDataStyleWithleftBorder", SectioncellDataStyleWithleftBorder);	
		
		CellStyle SectioncellDataStyleWithBottomBorder = workbook.createCellStyle();
		SectioncellDataStyleWithBottomBorder = workbook.createCellStyle();
		SectioncellDataStyleWithBottomBorder.setFont(headingCellDataFont);
		SectioncellDataStyleWithBottomBorder.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		SectioncellDataStyleWithBottomBorder.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		SectioncellDataStyleWithBottomBorder.setWrapText(true);
		SectioncellDataStyleWithBottomBorder.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		SectioncellDataStyleWithBottomBorder.setFillPattern(HSSFCellStyle.BORDER_NONE);
		/*SectioncellDataStyleWithBottomBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		SectioncellDataStyleWithBottomBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		SectioncellDataStyleWithBottomBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);*/
		SectioncellDataStyleWithBottomBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("SectioncellDataStyleWithBottomBorder", SectioncellDataStyleWithBottomBorder);
		
		CellStyle blackBackGroundWithWhiteFont = workbook.createCellStyle();
		blackBackGroundWithWhiteFont.setFont(cellDataFont1);
		blackBackGroundWithWhiteFont.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		blackBackGroundWithWhiteFont.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		blackBackGroundWithWhiteFont.setWrapText(true);
		blackBackGroundWithWhiteFont.setFillForegroundColor(getHSSFColor(HSSFColor.BLACK.index, 0, 0, 0));
		blackBackGroundWithWhiteFont.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		/*SectioncellDataStyleWithBottomBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		SectioncellDataStyleWithBottomBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		SectioncellDataStyleWithBottomBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);*/
		//blackBackGroundWithWhiteFont.setBorderTop(HSSFCellStyle.BORDER_THIN);
		blackBackGroundWithWhiteFont.setDataFormat(ch.createDataFormat().getFormat("$"+"##,###,###,###,##0"));
		excelStyleSet.setStyle("blackBackGroundWithWhiteFont", blackBackGroundWithWhiteFont);
		
		
		//////////////////
		//HeaderCaptionStyle///////////////////////////////////////
		
		
		Font headerCaptionFont = workbook.createFont();
		headerCaptionFont.setFontName("Areal");
		headerCaptionFont.setFontHeightInPoints((short) 10);
		headerCaptionFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerCaptionFont.setColor(HSSFColor.BLACK.index);
				
		
		CellStyle headerCaptionStyle = workbook.createCellStyle();
		headerCaptionStyle.setFont(headerCaptionFont);
		headerCaptionStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		headerCaptionStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle.setWrapText(true);
		headerCaptionStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		headerCaptionStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle", headerCaptionStyle);	
		//////////////////////////////////////
		
		
		//////////////////
		//CellDataHeadingStyle///////////////////////////////////////
		
		CellStyle cellDataHeadingStyle = workbook.createCellStyle();
		Font  cellDataHeadingFont= workbook.createFont();
		cellDataHeadingFont.setFontName("Areal");
		cellDataHeadingFont.setFontHeightInPoints((short) 12);
		cellDataHeadingFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		cellDataHeadingFont.setColor(HSSFColor.WHITE.index);
			
		
		cellDataHeadingStyle.setFont(cellDataHeadingFont);
		cellDataHeadingStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		cellDataHeadingStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		cellDataHeadingStyle.setWrapText(false);
		cellDataHeadingStyle.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_80_PERCENT.index, 192, 192, 192));
		cellDataHeadingStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellDataHeadingStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("cellDataHeadingStyle", cellDataHeadingStyle);	
		//////////////////////////////////////
		
		
		//////////////////
		//cellStyleWithNoDecimalStyle///////////////////////////////////////
		
		CellStyle cellStyleWithNoDecimalStyle  = workbook.createCellStyle();
		Font  cellStyleWithNoDecimalStyleFont= workbook.createFont();
		cellStyleWithNoDecimalStyleFont.setFontName("Areal");
		cellStyleWithNoDecimalStyleFont.setFontHeightInPoints((short) 10);
		cellStyleWithNoDecimalStyleFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		cellStyleWithNoDecimalStyleFont.setColor(HSSFColor.BLACK.index);
		
		
		cellStyleWithNoDecimalStyle.setFont(cellStyleWithNoDecimalStyleFont);
		cellStyleWithNoDecimalStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);	
		cellStyleWithNoDecimalStyle.setWrapText(false);
		cellStyleWithNoDecimalStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyleWithNoDecimalStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyleWithNoDecimalStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyleWithNoDecimalStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyleWithNoDecimalStyle.setWrapText(true);
		excelStyleSet.setStyle("cellStyleWithNoDecimalStyle", cellStyleWithNoDecimalStyle);	
		//////////////////////////////////////
		
		
		
		CellStyle NumberStyleNoCurrency  = workbook.createCellStyle();
		Font  columnH1Font= workbook.createFont();
		columnH1Font.setFontName("Areal");
		columnH1Font.setFontHeightInPoints((short) 10);
		columnH1Font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		columnH1Font.setColor(HSSFColor.BLACK.index);
		
		NumberStyleNoCurrency.setFont(columnH1Font);
		NumberStyleNoCurrency.setAlignment(HSSFCellStyle.ALIGN_CENTER);	
		NumberStyleNoCurrency.setWrapText(false);
		NumberStyleNoCurrency.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrency.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrency.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrency.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrency.setDataFormat(dataFormat.getFormat(this.getNumberFormat()));
		NumberStyleNoCurrency.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleNoCurrency", NumberStyleNoCurrency);	
		
		
		
		CellStyle CurrencyStyle  = workbook.createCellStyle();
		Font  CurrencyFont= workbook.createFont();
		CurrencyFont.setFontName("Areal");
		CurrencyFont.setFontHeightInPoints((short) 10);
		CurrencyFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		CurrencyFont.setColor(HSSFColor.BLACK.index);
		
		CurrencyStyle.setFont(CurrencyFont);
		CurrencyStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		CurrencyStyle.setWrapText(false);
		CurrencyStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		CurrencyStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		CurrencyStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		CurrencyStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		CurrencyStyle.setDataFormat(dataFormat.getFormat(currency_format));
		CurrencyStyle.setWrapText(true);
		excelStyleSet.setStyle("CurrencyStyle", CurrencyStyle);		
		
		
		
		//////////////////
		//cellDataHeadingStyle1///////////////////////////////////////
		
		Font cellDataHeadingFont1 = workbook.createFont();
		cellDataHeadingFont1.setFontName("Areal");
		cellDataHeadingFont1.setFontHeightInPoints((short) 10);
		cellDataHeadingFont1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		cellDataHeadingFont1.setColor(HSSFColor.WHITE.index);
		cellDataHeadingFont1.setUnderline(HSSFFont.U_SINGLE);

		CellStyle cellDataHeadingStyle1 = workbook.createCellStyle();
		cellDataHeadingStyle1.setFont(cellDataHeadingFont1);
		cellDataHeadingStyle1.setWrapText(true);
		cellDataHeadingStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellDataHeadingStyle1.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE.index, 52, 84, 156));
		cellDataHeadingStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellDataHeadingStyle1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("cellDataHeadingStyle1", cellDataHeadingStyle1);
		//////////////////////////////////////
		
		CellStyle cellDataWithNumberFormat_perscent_formatDecimal = workbook.createCellStyle();
		cellDataWithNumberFormat_perscent_formatDecimal.setFont(columnH1Font);
		cellDataWithNumberFormat_perscent_formatDecimal.setAlignment(HSSFCellStyle.ALIGN_CENTER);	
		cellDataWithNumberFormat_perscent_formatDecimal.setWrapText(false);
		cellDataWithNumberFormat_perscent_formatDecimal.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal.setDataFormat(dataFormat.getFormat(perscent_formatTwoDecimal));
		excelStyleSet.setStyle("cellDataWithNumberFormat_perscent_formatDecimal", cellDataWithNumberFormat_perscent_formatDecimal);

		//Contract
		Font headerCaptionFont_contract = workbook.createFont();
		headerCaptionFont.setFontName("Areal");
		headerCaptionFont.setFontHeightInPoints((short) 10);
		headerCaptionFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		headerCaptionFont.setColor(HSSFColor.BLACK.index);
				
		
		CellStyle headerCaptionStylecontract = workbook.createCellStyle();
		headerCaptionStylecontract.setFont(headerCaptionFont);
		headerCaptionStylecontract.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerCaptionStylecontract.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStylecontract.setWrapText(true);
		headerCaptionStylecontract.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		headerCaptionStylecontract.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStylecontract.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontract.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontract.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontract.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStylecontract", headerCaptionStylecontract);
		
		CellStyle headerCaptionStylecontractWithColorRed = workbook.createCellStyle();
		headerCaptionStylecontractWithColorRed.setFont(headerCaptionFont);
		headerCaptionStylecontractWithColorRed.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerCaptionStylecontractWithColorRed.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStylecontractWithColorRed.setWrapText(true);
		headerCaptionStylecontractWithColorRed.setFillForegroundColor(getHSSFColor(HSSFColor.RED.index, 252, 0, 4));
		headerCaptionStylecontractWithColorRed.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStylecontractWithColorRed.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorRed.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorRed.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorRed.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorRed.setDataFormat(ch.createDataFormat().getFormat(date_format_MMM_yy));
		excelStyleSet.setStyle("headerCaptionStylecontractWithColorRed", headerCaptionStylecontractWithColorRed);
		
		CellStyle headerCaptionStylecontractWithColorGreen = workbook.createCellStyle();
		headerCaptionStylecontractWithColorGreen.setFont(headerCaptionFont);
		headerCaptionStylecontractWithColorGreen.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerCaptionStylecontractWithColorGreen.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStylecontractWithColorGreen.setWrapText(true);
		headerCaptionStylecontractWithColorGreen.setFillForegroundColor(getHSSFColor(HSSFColor.GREEN.index, 5, 176, 78));
		headerCaptionStylecontractWithColorGreen.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStylecontractWithColorGreen.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorGreen.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorGreen.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorGreen.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headerCaptionStylecontractWithColorGreen.setDataFormat(ch.createDataFormat().getFormat(date_format_MMM_yy));
		excelStyleSet.setStyle("headerCaptionStylecontractWithColorGreen", headerCaptionStylecontractWithColorGreen);
		//////////////////////////////////////
		
		
		CellStyle cellDataWithNumberFormat_perscent_formatDecimal_contract  = workbook.createCellStyle();
		
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setFont(headerCaptionFont);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setWrapText(true);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellDataWithNumberFormat_perscent_formatDecimal_contract.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("cellDataWithNumberFormat_perscent_formatDecimal_contract", cellDataWithNumberFormat_perscent_formatDecimal_contract);
		
		
		CellStyle cellDataHeadingStyle_heading_contract1 = workbook.createCellStyle();
		cellDataHeadingStyle_heading_contract1.setFont(cellDataHeadingFont1);
		cellDataHeadingStyle_heading_contract1.setWrapText(true);
		cellDataHeadingStyle_heading_contract1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellDataHeadingStyle_heading_contract1.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_80_PERCENT.index, 70, 84, 107));
		cellDataHeadingStyle_heading_contract1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellDataHeadingStyle_heading_contract1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle_heading_contract1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle_heading_contract1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellDataHeadingStyle_heading_contract1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("cellDataHeadingStyle_heading_contract1", cellDataHeadingStyle_heading_contract1);
		
		
		Font headerCaptionFont_AuctionPreview_Tiered = workbook.createFont();
		headerCaptionFont_AuctionPreview_Tiered.setFontName("Areal");
		headerCaptionFont_AuctionPreview_Tiered.setFontHeightInPoints((short) 10);
		headerCaptionFont_AuctionPreview_Tiered.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerCaptionFont_AuctionPreview_Tiered.setColor(HSSFColor.BLACK.index);
		
		
		CellStyle headerCaptionStyle_AuctionPreview_Tiered = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_Tiered.setFont(headerCaptionFont_AuctionPreview_Tiered);
		headerCaptionStyle_AuctionPreview_Tiered.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerCaptionStyle_AuctionPreview_Tiered.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_Tiered.setWrapText(false);
		headerCaptionStyle_AuctionPreview_Tiered.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		headerCaptionStyle_AuctionPreview_Tiered.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle_AuctionPreview_Tiered.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_Tiered.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_Tiered.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_Tiered.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_Tiered", headerCaptionStyle_AuctionPreview_Tiered);	
		//////////////////////////////////////
		
		CellStyle headerCaptionStyle_AuctionPreview_TieredData = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_TieredData.setFont(headerCaptionFont_AuctionPreview_Tiered);
		headerCaptionStyle_AuctionPreview_TieredData.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerCaptionStyle_AuctionPreview_TieredData.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_TieredData.setWrapText(false);
		//headerCaptionStyle_AuctionPreview_TieredData.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		//headerCaptionStyle_AuctionPreview_TieredData.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle_AuctionPreview_TieredData.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_TieredData", headerCaptionStyle_AuctionPreview_TieredData);	
		//////////////////////////////////////
		
		CellStyle headerCaptionStyle_AuctionPreview_TieredHeading = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_TieredHeading.setFont(headerCaptionFont_AuctionPreview_Tiered);
		headerCaptionStyle_AuctionPreview_TieredHeading.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerCaptionStyle_AuctionPreview_TieredHeading.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_TieredHeading.setWrapText(false);
		headerCaptionStyle_AuctionPreview_TieredHeading.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		headerCaptionStyle_AuctionPreview_TieredHeading.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle_AuctionPreview_TieredHeading.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_TieredHeading", headerCaptionStyle_AuctionPreview_TieredHeading);	
		//////////////////////////////////////
		
		CellStyle headerCaptionStyle_AuctionPreview_TieredHeading_left = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setFont(headerCaptionFont_AuctionPreview_Tiered);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setWrapText(false);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_50_PERCENT.index, 191, 191, 191));
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_left.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_TieredHeading_left", headerCaptionStyle_AuctionPreview_TieredHeading_left);	
		//////////////////////////////////////
		
		CellStyle headerCaptionStyle_AuctionPreview_TieredHeading_lightGray = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setFont(headerCaptionFont_AuctionPreview_Tiered);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setWrapText(false);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 217, 217, 217));
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredHeading_lightGray.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_TieredHeading_lightGray", headerCaptionStyle_AuctionPreview_TieredHeading_lightGray);	
		//////////////////////////////////////
		
		
		Font headerCaptionStyle_AuctionPreview_Tiered_NormalFont = workbook.createFont();
		headerCaptionStyle_AuctionPreview_Tiered_NormalFont.setFontName("Areal");
		headerCaptionStyle_AuctionPreview_Tiered_NormalFont.setFontHeightInPoints((short) 10);
		headerCaptionStyle_AuctionPreview_Tiered_NormalFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		headerCaptionStyle_AuctionPreview_Tiered_NormalFont.setColor(HSSFColor.BLACK.index);
		
		CellStyle headerCaptionStyle_AuctionPreview_TieredData_NormalFont = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setFont(headerCaptionStyle_AuctionPreview_Tiered_NormalFont);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setWrapText(false);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerCaptionStyle_AuctionPreview_TieredData_NormalFont.setBorderTop(HSSFCellStyle.BORDER_THIN);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_TieredData_NormalFont", headerCaptionStyle_AuctionPreview_TieredData_NormalFont);	
		
		Font headerCaptionStyle_AuctionPreview_Tiered_HeadingFont = workbook.createFont();
		headerCaptionStyle_AuctionPreview_Tiered_HeadingFont.setFontName("Areal");
		headerCaptionStyle_AuctionPreview_Tiered_HeadingFont.setFontHeightInPoints((short) 14);
		headerCaptionStyle_AuctionPreview_Tiered_HeadingFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerCaptionStyle_AuctionPreview_Tiered_HeadingFont.setColor(HSSFColor.BLACK.index);
		
		CellStyle headerCaptionStyle_AuctionPreview_Tiered_Heading = workbook.createCellStyle();
		headerCaptionStyle_AuctionPreview_Tiered_Heading.setFont(headerCaptionStyle_AuctionPreview_Tiered_HeadingFont);
		headerCaptionStyle_AuctionPreview_Tiered_Heading.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		excelStyleSet.setStyle("headerCaptionStyle_AuctionPreview_Tiered_Heading", headerCaptionStyle_AuctionPreview_Tiered_Heading);	
		
		Font headerCaptionStyle_Contract_HeadingFont = workbook.createFont();
		headerCaptionStyle_Contract_HeadingFont.setFontName("Arial");
		headerCaptionStyle_Contract_HeadingFont.setFontHeightInPoints((short) 12);
		headerCaptionStyle_Contract_HeadingFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerCaptionStyle_Contract_HeadingFont.setColor(HSSFColor.BLACK.index);
		
		CellStyle headerCaptionStyle_Contract_Heading = workbook.createCellStyle();
		headerCaptionStyle_Contract_Heading.setFont(headerCaptionStyle_Contract_HeadingFont);
		headerCaptionStyle_Contract_Heading.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		excelStyleSet.setStyle("headerCaptionStyle_Contract_Heading", headerCaptionStyle_Contract_Heading);
		
		CellStyle Contract_numberFormat = workbook.createCellStyle();
		Contract_numberFormat.setFont(headerCaptionFont);
		Contract_numberFormat.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		Contract_numberFormat.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		Contract_numberFormat.setWrapText(true);
		Contract_numberFormat.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		Contract_numberFormat.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Contract_numberFormat.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormat.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormat.setBorderRight(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormat.setBorderTop(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormat.setDataFormat(ch.createDataFormat().getFormat(number_format3));
		excelStyleSet.setStyle("Contract_numberFormat", Contract_numberFormat);
		
		CellStyle Contract_numberFormatComma = workbook.createCellStyle();
		Contract_numberFormatComma.setFont(headerCaptionFont);
		Contract_numberFormatComma.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		Contract_numberFormatComma.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		Contract_numberFormatComma.setWrapText(true);
		Contract_numberFormatComma.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		Contract_numberFormatComma.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Contract_numberFormatComma.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatComma.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatComma.setBorderRight(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatComma.setBorderTop(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatComma.setDataFormat(ch.createDataFormat().getFormat(number_format2));
		excelStyleSet.setStyle("Contract_numberFormatComma", Contract_numberFormatComma);
		
		CellStyle Contract_numberFormatCommadecimal = workbook.createCellStyle();
		Contract_numberFormatCommadecimal.setFont(headerCaptionFont);
		Contract_numberFormatCommadecimal.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		Contract_numberFormatCommadecimal.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		Contract_numberFormatCommadecimal.setWrapText(true);
		Contract_numberFormatCommadecimal.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		Contract_numberFormatCommadecimal.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Contract_numberFormatCommadecimal.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatCommadecimal.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatCommadecimal.setBorderRight(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatCommadecimal.setBorderTop(HSSFCellStyle.BORDER_THIN);
		Contract_numberFormatCommadecimal.setDataFormat(ch.createDataFormat().getFormat(number_format1));
		excelStyleSet.setStyle("Contract_numberFormatCommadecimal", Contract_numberFormatCommadecimal);
		
		CellStyle Contract_DateMMM_YY = workbook.createCellStyle();
		Contract_DateMMM_YY.setFont(headerCaptionFont);
		Contract_DateMMM_YY.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		Contract_DateMMM_YY.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);
		Contract_DateMMM_YY.setWrapText(true);
		Contract_DateMMM_YY.setFillForegroundColor(getHSSFColor(HSSFColor.WHITE.index, 255, 255, 255));
		Contract_DateMMM_YY.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Contract_DateMMM_YY.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		Contract_DateMMM_YY.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		Contract_DateMMM_YY.setBorderRight(HSSFCellStyle.BORDER_THIN);
		Contract_DateMMM_YY.setBorderTop(HSSFCellStyle.BORDER_THIN);
		Contract_DateMMM_YY.setDataFormat(ch.createDataFormat().getFormat(date_format_MMM_yy));
		excelStyleSet.setStyle("Contract_DateMMM_YY", Contract_DateMMM_YY);
		

		CellStyle percentageOddDataStyle_AICP = workbook.createCellStyle();
		percentageOddDataStyle_AICP.setFont(columnFont);
		percentageOddDataStyle_AICP.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		percentageOddDataStyle_AICP.setWrapText(false);
		percentageOddDataStyle_AICP.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		percentageOddDataStyle_AICP.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		percentageOddDataStyle_AICP.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle_AICP.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle_AICP.setBorderRight(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle_AICP.setBorderTop(HSSFCellStyle.BORDER_THIN);
		percentageOddDataStyle_AICP.setDataFormat(dataFormat.getFormat(percentage_format));
		percentageOddDataStyle_AICP.setWrapText(true);
		excelStyleSet.setStyle("percentageOddDataStyle_AICP", percentageOddDataStyle_AICP);	
		
		CellStyle percentageEvenDataStyle_AICP  = workbook.createCellStyle();
		percentageEvenDataStyle_AICP.setFont(columnFont);
		percentageEvenDataStyle_AICP.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		percentageEvenDataStyle_AICP.setWrapText(false);
		percentageEvenDataStyle_AICP.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		percentageEvenDataStyle_AICP.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		percentageEvenDataStyle_AICP.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle_AICP.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle_AICP.setBorderRight(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle_AICP.setBorderTop(HSSFCellStyle.BORDER_THIN);
		percentageEvenDataStyle_AICP.setDataFormat(dataFormat.getFormat(percentage_format));
		percentageEvenDataStyle_AICP.setWrapText(true);
		excelStyleSet.setStyle("percentageEvenDataStyle_AICP", percentageEvenDataStyle_AICP);	
		
		
		CellStyle NumberStyleNoCurrencyOdd_AICP  = workbook.createCellStyle();
		NumberStyleNoCurrencyOdd_AICP.setFont(columnFont);
		NumberStyleNoCurrencyOdd_AICP.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleNoCurrencyOdd_AICP.setWrapText(false);
		NumberStyleNoCurrencyOdd_AICP.setFillForegroundColor(getHSSFColor(HSSFColor.GREY_25_PERCENT.index, 200, 200, 200));
		NumberStyleNoCurrencyOdd_AICP.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleNoCurrencyOdd_AICP.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd_AICP.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd_AICP.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd_AICP.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyOdd_AICP.setDataFormat(ch.createDataFormat().getFormat("##,###,###,###,##0"));
		NumberStyleNoCurrencyOdd_AICP.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleNoCurrencyOdd_AICP", NumberStyleNoCurrencyOdd_AICP);	
		
		CellStyle NumberStyleNoCurrencyEven_AICP  = workbook.createCellStyle();
		NumberStyleNoCurrencyEven_AICP.setFont(columnFont);
		NumberStyleNoCurrencyEven_AICP.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
		NumberStyleNoCurrencyEven_AICP.setWrapText(false);
		NumberStyleNoCurrencyEven_AICP.setFillForegroundColor(getHSSFColor(HSSFColor.BLUE_GREY.index, 244, 244, 244));
		NumberStyleNoCurrencyEven_AICP.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		NumberStyleNoCurrencyEven_AICP.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven_AICP.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven_AICP.setBorderRight(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven_AICP.setBorderTop(HSSFCellStyle.BORDER_THIN);
		NumberStyleNoCurrencyEven_AICP.setDataFormat(ch.createDataFormat().getFormat("##,###,###,###,##0"));
		NumberStyleNoCurrencyEven_AICP.setWrapText(true);
		excelStyleSet.setStyle("NumberStyleNoCurrencyEven_AICP", NumberStyleNoCurrencyEven_AICP);	
	}
}
