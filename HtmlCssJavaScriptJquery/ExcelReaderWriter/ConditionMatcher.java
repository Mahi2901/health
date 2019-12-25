package excelreader;

import org.apache.poi.ss.usermodel.Cell;

public abstract class ConditionMatcher
{
	Cell		cell;
	int			currentRowIndex		= 0;
	int			currentColumnIndex	= 0;
	CellData	matchCellData		= null;
	boolean		isMatch				= false;

	public void setMatch(boolean isMatch)
	{
		this.isMatch = isMatch;
	}

	public boolean isMatch()
	{
		return isMatch;
	}

	public Cell getCell()
	{
		return cell;
	}

	public void setCell(Cell cell)
	{
		this.cell = cell;
	}

	public int getCurrentRowIndex()
	{
		return currentRowIndex;
	}

	public void setCurrentRowIndex(int currentRowIndex)
	{
		this.currentRowIndex = currentRowIndex;
	}

	public int getCurrentColumnIndex()
	{
		return currentColumnIndex;
	}

	public void setCurrentColumnIndex(int currentColumnIndex)
	{
		this.currentColumnIndex = currentColumnIndex;
	}

	public CellData getMatchCellData()
	{
		return matchCellData;
	}

	public void setMatchCellData(CellData matchCellData)
	{
		this.matchCellData = matchCellData;
	}
	public void setDefault()
	{
		cell = null;
		currentRowIndex = 0;
		currentColumnIndex = 0;
		matchCellData = null;
		isMatch = false;
	}

	public abstract boolean match(CellData excelCellData);

	public void print()
	{
		System.out.println("_________ConditionMatcher________");
		System.out.println(":->isMatch            =" + isMatch);
		System.out.println(":->currentRowIndex    =" + currentRowIndex);
		System.out.println(":->currentColumnIndex =" + currentColumnIndex);
		System.out.println(":->cellis null?       =" + (cell == null));
	}
}
