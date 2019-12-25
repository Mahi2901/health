package excelreader;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelStyleSet
{
	private HashMap<String, CellStyle> styleHashMap;
	public ExcelStyleSet()
	{
		styleHashMap = new HashMap<String, CellStyle>(2);
	}
	public void setStyle(String key, CellStyle cellStyle)
	{
		styleHashMap.put(key, cellStyle);
	}
	public CellStyle getStyle(String key)
	{
		CellStyle cellStyle = null;
		cellStyle = styleHashMap.get(key);
		return cellStyle;
		
	}
	/*public void print(String preFix)
	{
		if (styleHashMap != null)
		{
			System.out.println(preFix + ",styleHashMap.size()===" + styleHashMap.size());
		}
		else
		{
			System.out.println(preFix + ",styleHashMap.size()===null-0");
		}
	}*/
}
