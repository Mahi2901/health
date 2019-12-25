package excelreader;

import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

public class CellData
{
	public CellType	cellType;
	public boolean	isCellNull			= false;
	public boolean	haveValue			= false;
	public boolean	isNone;
	public boolean	isBlank;
	public Boolean	booleanData;
	public byte		errorCellValue;
	public String	formulaString;
	public double	numeric;
	public String	stringData;
	public String	debugPoint			= "";
	public String	cellReference		= "";
	public int		currentColumnIndex	= 0;
	public int		currentRowIndex		= 0;

	public void setDefault()
	{
		isNone = false;
		isBlank = false;
		haveValue = false;
		booleanData = null;
		errorCellValue = 0;
		formulaString = null;
		numeric = 0;
		stringData = null;
		errorCellValue = 0;
	}

	public double getNumeric(double defaultValue)
	{
		if (cellType == CellType.NUMERIC)
		{
			if (this.haveValue == true)
			{
				return numeric;
			}
		}
		return defaultValue;
	}

	public String getStringData()
	{
		if (cellType == CellType.STRING)
		{
			if (this.haveValue == true)
				return stringData;
		}
		return null;
	}
	public Date getDateXX()
	{
		Date javaDate = null;
		double tmp = getNumeric(0);
		if (tmp > 0)
		{
			javaDate = DateUtil.getJavaDate(tmp);
		}
		return javaDate;
	}
	public Date getDate()
	{
		Date javaDate = null;
		if (numeric > 0)
		{
			javaDate = DateUtil.getJavaDate(numeric);
		}
		return javaDate;
	}
	private boolean haveValueNotUsed()
	{
		boolean proper = false;
		//ERROR  errorCellValue
		if (!(isCellNull || isBlank || isNone || errorCellValue != 0))
		{
			switch (this.cellType)
			{
				case BOOLEAN:
					if (this.booleanData != null)
					{
						proper = true;
					}
					break;
				case NUMERIC:
					if (this.numeric != 0)
					{
						proper = true;
					}
					break;
				case STRING:
					if (this.stringData != null)
					{
						proper = true;
					}
					break;
				default:
					System.out.print("inside the default..Cellread :: case-198");
			}

		}
		return proper;
	}
	public void print()
	{
		System.out.println("_______________CellData_____________" + debugPoint);
		System.out.println(":->isCellNull;            " + isCellNull);
		System.out.println(":->isNone;            " + isNone);
		System.out.println(":->isBlank;           " + isBlank);
		System.out.println(":->booleanData;       " + booleanData);
		System.out.println(":->errorCellValue;    " + errorCellValue);
		System.out.println(":->formulaString;     " + formulaString);
		System.out.println(":->numeric;           " + numeric);
		System.out.println(":->stringData;        " + stringData);

	}
}
