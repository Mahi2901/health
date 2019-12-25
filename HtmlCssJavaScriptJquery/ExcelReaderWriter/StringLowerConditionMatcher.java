package excelreader;

import org.apache.poi.ss.usermodel.CellType;

public class StringLowerConditionMatcher extends ConditionMatcher
{

	String matchWith;

	public String getMatchWith()
	{
		return matchWith;
	}

	public void setMatchWith(String matchWith)
	{
		this.matchWith = matchWith;
	}

	@Override
	public boolean match(CellData excelCellData)
	{
		if (excelCellData.stringData != null)
		{
			if (excelCellData.cellType == CellType.STRING || excelCellData.cellType == CellType.FORMULA)
			{
				return excelCellData.stringData.trim().toLowerCase().equals(matchWith);
			}

		}
		return false;
	}
}
