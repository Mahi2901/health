package excelreader;

import java.io.IOException;

import com.sun.rowset.internal.Row;

public class ExcelReader extends CellReader
{
	public BookmarkCellPointer	bookmarkCellPointer;
	CellData					cellData[]	= null;

	public ExcelReader(String fileNameWithPath) throws Exception
	{
		super(fileNameWithPath);
		bookmarkCellPointer = new BookmarkCellPointer();
	}
	public void bookmark(String bookmarkName)
	{
		bookmark(bookmarkName, bookmarkCellPointer);
	}
	public void setCellAddressByBookmark(String bookmarkName)
	{
		setCellAddress(bookmarkName, bookmarkCellPointer);
	}
	public void newMaxBufferDataCell(int arrayLength)
	{
		if (this.cellData != null)
		{
			if (this.cellData.length != arrayLength)
			{
				this.cellData = null;
				this.cellData = new CellData[arrayLength];
			}
		}
		else
		{
			this.cellData = new CellData[arrayLength];
		}
		for (int i = 0; i < cellData.length; i++)
		{
			cellData[i] = new CellData();
		}

	}
	//[][][][][][]==false If value found at one place return true
	public boolean anyOneCellArrayHaveValue(int numberOfColumns)
	{
		//int numberOfColumns = cellData.length;
		for (int i = 0; i < numberOfColumns; i++)
		{
			if (cellData[i].haveValue)
			{
				return true;
			}
		}
		return false;
	}
	public void readRow(int numberOfColumns)
	{
		for (int i = 0; i < numberOfColumns; i++)
		{
			setCellData(cellData[i]);
			cellData[i] = read();
			move(0, 1);
		}

		printReadRow(numberOfColumns);
	}
	public void readColumn(int numberOfRows)
	{

		for (int i = 0; i < numberOfRows; i++)
		{
			setCellData(cellData[i]);
			cellData[i] = read();
			move(1, 0);
		}
		printReadRow(numberOfRows);
	}
	public void printReadRow(int numberOfColumns)
	{
		for (int i = 0; i < numberOfColumns; i++)
		{
			cellData[i].print();
		}
	}
	public int findEmptyDown(ConditionMatcher conditionMatcher, int uptoNumberOfRows, int allowEmptyCount)
	{
		int lastNotFoundAt = -1;
		int emptyRowCounter = 0;
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		for (int i = 0; i < uptoNumberOfRows; i++)
		{

			cellData = read();
			if (conditionMatcher.match(cellData))
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				lastNotFoundAt = i + 1;
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			move(1, 0);

		}
		///////////////////////
		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
		return lastNotFoundAt;

	}
	public int findEmptyDown(int uptoNumberOfRows, int allowEmptyCount)
	{
		int lastNotFoundAt = -1;
		int emptyRowCounter = 0;
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		for (int i = 0; i < uptoNumberOfRows; i++)
		{
			cellData = read();
			if (cellData.isCellNull)
			{
				emptyRowCounter++;
			}
			else if (cellData.isBlank || cellData.isNone)
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				lastNotFoundAt = i + 1;
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			move(1, 0);
		}
		///////////////////////
		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
		return lastNotFoundAt;

	}
	public int findEmptyRight(int uptoNumberOfColumns, int allowEmptyCount)
	{
		int lastNotFoundAt = -1;
		int emptyRowCounter = 0;
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		for (int i = 0; i < uptoNumberOfColumns; i++)
		{

			cellData = read();
			if (cellData.isCellNull)
			{
				emptyRowCounter++;
			}
			else if (cellData.isBlank || cellData.isNone)
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				lastNotFoundAt = i + 1;
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			move(0, 1);

		}
		///////////////////////
		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
		return lastNotFoundAt;
	}

	public String findEmptyRightCellAddress(int uptoNumberOfColumns, int allowEmptyCount)
	{
		int lastNotFoundAt = -1;
		int emptyRowCounter = 0;
		String cellAddress = null;

		//////////////////////
		CellData cellData = null;
		for (int i = 0; i < uptoNumberOfColumns; i++)
		{

			cellData = read();
			cellData.print();
			if (cellData.isCellNull)
			{
				emptyRowCounter++;
			}
			else if (cellData.isBlank || cellData.isNone)
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				lastNotFoundAt = i + 1;
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			else
			{
				move(0, 1);
			}

		}
		///////////////////////
		cellAddress = cellData.cellReference;
		return cellAddress;
	}
	public int findEmptyRight(ConditionMatcher conditionMatcher, int uptoNumberOfColumns, int allowEmptyCount)
	{
		int lastNotFoundAt = -1;
		int emptyRowCounter = 0;
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		for (int i = 0; i < uptoNumberOfColumns; i++)
		{

			cellData = read();
			if (conditionMatcher.match(cellData))
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				lastNotFoundAt = i + 1;
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			move(0, 1);
		}
		///////////////////////
		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
		return lastNotFoundAt;
	}
	public int findUntillTextDown(String text, int uptoNumberOfRows, int allowEmptyCount)
	{
		int lastFoundAt = -1;
		int emptyRowCounter = 0;
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		for (int i = 0; i < uptoNumberOfRows; i++)
		{
			cellData = read();
			if (cellData.isCellNull)
			{
				emptyRowCounter++;
			}
			else if (cellData.isBlank || cellData.isNone)
			{
				emptyRowCounter++;
			}
			else
			{
				emptyRowCounter = 0;
				if (cellData.stringData.indexOf(text) >= 0)
				{
					lastFoundAt = i + 1;
				}
				else
				{
					break;
				}
			}
			if (emptyRowCounter >= allowEmptyCount)
			{
				break;
			}
			move(1, 0);
		}
		///////////////////////
		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
		return lastFoundAt;
	}
	public void find(ConditionMatcher conditionMatcher, int uptoNumberOfRows, int uptoNumberOfColumns)
	{
		find(conditionMatcher, this.currentRowIndex, this.getCurrentColumnIndex(), uptoNumberOfRows, uptoNumberOfColumns);
	}
	public void find(ConditionMatcher conditionMatcher, int startRow, int startColumn, int uptoNumberOfRows, int uptoNumberOfColumns)
	{
		conditionMatcher.setDefault();
		int[] bookamarkCellIndex = getCellIndex();
		//////////////////////
		CellData cellData;
		int rowIndex = startRow;
		int colIndex = startColumn;
		int rowIndex_upto = startRow + uptoNumberOfRows;
		int colIndex_upto = startColumn + uptoNumberOfColumns;
		setCellAddress(startRow, startColumn);
		for (; rowIndex < rowIndex_upto; rowIndex++)
		{
			setCellAddress(rowIndex, startColumn);
			for (colIndex = startColumn; colIndex < colIndex_upto; colIndex++)
			{
				setCellAddress(rowIndex, colIndex);
				cellData = read();
				cellData.print();

				if (conditionMatcher.match(cellData))//if (cellData.cellType == CellType.STRING)
				{
					conditionMatcher.setMatch(true);
					conditionMatcher.setCell(getCell());
					conditionMatcher.setCurrentRowIndex(rowIndex);
					conditionMatcher.setCurrentColumnIndex(colIndex);
					///////////////////////

					setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
					return;
				}
			}
		}
		///////////////////////

		setCellAddress(bookamarkCellIndex[0], bookamarkCellIndex[1]);
	}

	public CellData[] getCellDataArray()
	{
		return cellData;
	}

	public void close() throws IOException
	{
		super.close();
	}


}
