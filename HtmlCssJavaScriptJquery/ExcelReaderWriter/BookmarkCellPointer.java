package excelreader;

import java.util.ArrayList;
import java.util.HashMap;

public class BookmarkCellPointer
{
	HashMap<String, int[]> pointerHashMap = new HashMap<String, int[]>();
	public BookmarkCellPointer()
	{
		pointerHashMap = new HashMap<String, int[]>(3);
	}
	public void mark(String bookmarkName, int[] pointer)
	{
		pointerHashMap.put(bookmarkName, pointer);
	}
	public int[] get(String bookmarkName)
	{
		return pointerHashMap.get(bookmarkName);
	}
}
