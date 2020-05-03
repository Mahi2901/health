public class FormBase extends ValidatorForm implements java.io.Serializable
{

  private String			rowKey;
	private String			pageNo;
	private String			recordsPerPage;
  private String			orderBy					= "1";
  private String			searchBy;
	public String 			lmSelectedMenu="0";
	public String       lmSelectedSubMenu="0";
  
  public String getRowKey()
	{
		return rowKey;
	}

	public void setRowKey(String rowKey)
	{
		this.rowKey = rowKey;
	}
  
  public String getPageNo()
	{
			return this.pageNo;
	}

	public void setPageNo(String pageNo)
	{
		this.pageNo = pageNo;
	}
  public String getRecordsPerPage()
	{
			return this.recordsPerPage;
	}

	public void setRecordsPerPage(String recordsPerPage)
	{
		this.recordsPerPage = recordsPerPage;
	}
  
  public String getOrderBy()
	{
			return this.orderBy;
	}

	public void setOrderBy(String orderBy)
	{
		this.orderBy = orderBy;
	}
  
  public String getSearchBy()
	{
		return searchBy;
	}

	public void setSearchBy(String searchBy)
	{
		this.searchBy = searchBy;
	}
  
  public String getLmSelectedMenu() {
		return lmSelectedMenu;
	}
	public void setLmSelectedMenu(String lmSelectedMenu) {
		this.lmSelectedMenu = lmSelectedMenu;
	}
	public String getLmSelectedSubMenu() {
		return lmSelectedSubMenu;
	}
	public void setLmSelectedSubMenu(String lmSelectedSubMenu) {
		this.lmSelectedSubMenu = lmSelectedSubMenu;
	}
  
  public static String showCommonListHidden(HttpServletRequest request)
	{
		String hiddenvar = "";

		//List Paras...
		if (null != request.getParameter("rowKey"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"rowKey\" id=\"rowKey\" value=\"" + General.chkHTMLStr(request.getParameter("rowKey")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"rowKey\" value=\""+General.chkHTMLStr("-1")+"\" >";
		if (null != request.getParameter("pageNo"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"pageNo\" value=\"" + General.chkHTMLStr(request.getParameter("pageNo")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"pageNo\" value=\""+General.chkHTMLStr("1")+"\" >";

		if (null != request.getParameter("recordsPerPage"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"recordsPerPage\" value=\"" + General.chkHTMLStr(request.getParameter("recordsPerPage")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"recordsPerPage\" value=\""+General.chkHTMLStr("10")+"\" >";
		
    if (null != request.getParameter("orderBy"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"orderBy\" value=\"" +General.chkHTMLStr( request.getParameter("orderBy")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"orderBy\" value=\""+General.chkHTMLStr("1")+"\" >";
      
		if (null != request.getParameter("lmSelectedMenu"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"lmSelectedMenu\" value=\"" + General.chkHTMLStr( request.getParameter("lmSelectedMenu")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"lmSelectedMenu\" value=\""+General.chkHTMLStr("0")+"\" >";
		
		if (null != request.getParameter("lmSelectedSubMenu"))
			hiddenvar += "\n\t<input type=\"hidden\" name=\"lmSelectedSubMenu\" value=\"" + General.chkHTMLStr( request.getParameter("lmSelectedSubMenu")) + "\" >";
		else
			hiddenvar += "\n\t<input type=\"hidden\" name=\"lmSelectedSubMenu\" value=\""+General.chkHTMLStr("0")+"\" >";
		
		return hiddenvar;
	}
  
  public static String chkHTMLStr(String str)
	{
		String str1 = "";
		if (str == null)
			return str1;
		
		for (int i = 0; i < str.length(); i++)
		{
			switch (str.charAt(i))
			{
				case '"':
					str1 += "&quot;";
					break;
				case '<':
					str1 += "&lt;";
					break;
				case '>':
					str1 += "&gt;";
					break;
				case '\t':
					str1 += "&nbsp;&nbsp;&nbsp;";
					break;
				default:
					str1 += str.charAt(i);
					break;
			}
		}
		return str1;
	}
  
  public static String chkXMLStr(String str)
	{
		String str1 = "";
		if (str == null)
			return str1;

		str1 = str.replaceAll("&", "&amp;");
		str1 = str1.replaceAll("<", "&lt;");
		str1 = str1.replaceAll(">", "&gt;");
		// str = str.replaceAll("\\", "\\\\");
		str1 = str1.replaceAll("\"", "&quot;");
		str1 = str1.replaceAll("'", "&apos;");

		return str1;
	}
  
}

<%-- 
<%
	String pgInfoData="";
	pgInfoData+="\n\n"+(this.getClass().getName()).replace("org.apache.jsp.","D:\\WS\\npp\\nppWeb\\WebContent\\").replaceAll("\\.","\\\\").replace("_jsp",".jsp");
	out.print("<input type='hidden' value='"+pgInfoData.trim()+"'/>");
%>
--%>
