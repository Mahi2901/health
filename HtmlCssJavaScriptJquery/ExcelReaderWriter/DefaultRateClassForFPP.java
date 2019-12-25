package excelreader;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;

import excelreader.ExcelReader;

public class DefaultRateClassForFPP
{

	ArrayList<DefaultRateClassZone> defaultRateClassZonesList = null;
	public void setDeafultValue()
	{
		defaultRateClassZonesList = new ArrayList<DefaultRateClassZone>();

		DefaultRateClassZone defaultRateClassZone1 = new DefaultRateClassZone();
		defaultRateClassZone1.setUtilityName("National Grid - RI");
		defaultRateClassZone1.setAbbreviation("RhodeIsland");
		defaultRateClassZone1.setSupplierType(2);
		defaultRateClassZone1.setRateClassName("A16");
		defaultRateClassZonesList.add(defaultRateClassZone1);

		DefaultRateClassZone defaultRateClassZone2 = new DefaultRateClassZone();
    	defaultRateClassZone2.setUtilityName("Unitil - NH");
		defaultRateClassZone2.setAbbreviation("NewHampshire");
		defaultRateClassZone2.setRateClassName("Res(10,13,A)");
		defaultRateClassZone2.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone2);

//		DefaultRateClassZone defaultRateClassZone3 = new DefaultRateClassZone();
//		defaultRateClassZone3.setUtilityName("Eversource - MA (NSTAR)");
//		defaultRateClassZone3.setAbbreviation("Massachusetts");
//		defaultRateClassZone3.setRateClassName("Commercial");
//		defaultRateClassZone3.setSupplierType("Commercial");
//		defaultRateClassZonesList.add(defaultRateClassZone3);
//
//		DefaultRateClassZone defaultRateClassZone4 = new DefaultRateClassZone();
//		defaultRateClassZone4.setUtilityName("Eversource - MA (NSTAR)");
//		defaultRateClassZone4.setAbbreviation("Massachusetts");
//		defaultRateClassZone4.setRateClassName("Commercial");
//		defaultRateClassZone4.setSupplierType("Commercial");
//		defaultRateClassZonesList.add(defaultRateClassZone4);

		DefaultRateClassZone defaultRateClassZone5 = new DefaultRateClassZone();
		defaultRateClassZone5.setUtilityName("Eversource - MA (NSTAR)");
		defaultRateClassZone5.setAbbreviation("Massachusetts");
		defaultRateClassZone5.setRateClassName("Commercial");
		defaultRateClassZone5.setSupplierType(1);
		defaultRateClassZonesList.add(defaultRateClassZone5);

//		DefaultRateClassZone defaultRateClassZone6 = new DefaultRateClassZone();
//		defaultRateClassZone6.setUtilityName("Eversource - MA (NSTAR)");
//		defaultRateClassZone6.setAbbreviation("Massachusetts");
//		defaultRateClassZone6.setRateClassName("Residential");
//		defaultRateClassZone6.setSupplierType("Residential");
//		defaultRateClassZonesList.add(defaultRateClassZone6);
//
//		DefaultRateClassZone defaultRateClassZone8 = new DefaultRateClassZone();
//		defaultRateClassZone8.setUtilityName("Eversource - MA (NSTAR)");
//		defaultRateClassZone8.setAbbreviation("Massachusetts");
//		defaultRateClassZone8.setRateClassName("Residential");
//		defaultRateClassZone8.setSupplierType("Residential");
//		defaultRateClassZonesList.add(defaultRateClassZone8);

		DefaultRateClassZone defaultRateClassZone9 = new DefaultRateClassZone();
		defaultRateClassZone9.setUtilityName("Eversource - MA (NSTAR)");
		defaultRateClassZone9.setAbbreviation("Massachusetts");
		defaultRateClassZone9.setRateClassName("Residential");
		defaultRateClassZone9.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone9);

		DefaultRateClassZone defaultRateClassZone10 = new DefaultRateClassZone();
		defaultRateClassZone10.setUtilityName("Eversource - MA (WMECO)");
		defaultRateClassZone10.setAbbreviation("Massachusetts");
		defaultRateClassZone10.setRateClassName("R");
		defaultRateClassZone10.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone10);

		DefaultRateClassZone defaultRateClassZone11 = new DefaultRateClassZone();
		defaultRateClassZone11.setUtilityName("Eversource - NH (PSNH)");
		defaultRateClassZone11.setAbbreviation("NewHampshire");
		defaultRateClassZone11.setRateClassName("R");
		defaultRateClassZone11.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone11);

		DefaultRateClassZone defaultRateClassZone12 = new DefaultRateClassZone();
		defaultRateClassZone12.setUtilityName("Liberty Utilities - NH");
		defaultRateClassZone12.setAbbreviation("NewHampshire");
		defaultRateClassZone12.setRateClassName("R");
		defaultRateClassZone12.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone12);

		DefaultRateClassZone defaultRateClassZone13 = new DefaultRateClassZone();
		defaultRateClassZone13.setUtilityName("National Grid - MA");
		defaultRateClassZone13.setAbbreviation("Massachusetts");
		defaultRateClassZone13.setRateClassName("R");
		defaultRateClassZone13.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone13);

		DefaultRateClassZone defaultRateClassZone14 = new DefaultRateClassZone();
		defaultRateClassZone14.setUtilityName("Delmarva - DE");
		defaultRateClassZone14.setAbbreviation("Delaware");
		defaultRateClassZone14.setRateClassName("Residential");
		defaultRateClassZone14.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone14);

		DefaultRateClassZone defaultRateClassZone15 = new DefaultRateClassZone();
		defaultRateClassZone15.setUtilityName("CMP - ME");
		defaultRateClassZone15.setAbbreviation("Maine");
		defaultRateClassZone15.setRateClassName("Residential");
		defaultRateClassZone15.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone15);

		DefaultRateClassZone defaultRateClassZone16 = new DefaultRateClassZone();
		defaultRateClassZone16.setUtilityName("FGE - MA (Unitil)");
		defaultRateClassZone16.setAbbreviation("Massachusetts");
		defaultRateClassZone16.setRateClassName("R");
		defaultRateClassZone16.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone16);

		DefaultRateClassZone defaultRateClassZone17 = new DefaultRateClassZone();
		defaultRateClassZone17.setUtilityName("PSEG - NJ");
		defaultRateClassZone17.setAbbreviation("NewJersey");
		defaultRateClassZone17.setRateClassName("Residential");
		defaultRateClassZone17.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone17);

		DefaultRateClassZone defaultRateClassZone18 = new DefaultRateClassZone();
		defaultRateClassZone18.setUtilityName("JCP&L - NJ");
		defaultRateClassZone18.setAbbreviation("NewJersey");
		defaultRateClassZone18.setRateClassName("Residential");
		defaultRateClassZone18.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone18);

		DefaultRateClassZone defaultRateClassZone19 = new DefaultRateClassZone();
		defaultRateClassZone19.setUtilityName("Atlantic City Electric - NJ");
		defaultRateClassZone19.setAbbreviation("NewJersey");
		defaultRateClassZone19.setRateClassName("Residential");
		defaultRateClassZone19.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone19);

		DefaultRateClassZone defaultRateClassZone20 = new DefaultRateClassZone();
		defaultRateClassZone20.setUtilityName("NHEC - NH");
		defaultRateClassZone20.setAbbreviation("NewHampshire");
		defaultRateClassZone20.setRateClassName("Commercial");
		defaultRateClassZone20.setSupplierType(1);
		defaultRateClassZonesList.add(defaultRateClassZone20);

		DefaultRateClassZone defaultRateClassZone21 = new DefaultRateClassZone();
		defaultRateClassZone21.setUtilityName("NHEC - NH");
		defaultRateClassZone21.setAbbreviation("NewHampshire");
		defaultRateClassZone21.setRateClassName("Residential");
		defaultRateClassZone21.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone21);

		DefaultRateClassZone defaultRateClassZone22 = new DefaultRateClassZone();
		defaultRateClassZone22.setUtilityName("BGE - MD");
		defaultRateClassZone22.setAbbreviation("Maryland");
		defaultRateClassZone22.setRateClassName("R");
		defaultRateClassZone22.setSupplierType(2);
		defaultRateClassZonesList.add(defaultRateClassZone22);
	}

	public boolean isAllRateClass(String abbrevation, String utilitytName, String rateClassName)
	{
		boolean isAllRateClass = false;
		for (Iterator iterator = defaultRateClassZonesList.iterator(); iterator.hasNext();)
		{
			DefaultRateClassZone defaultRateClassZone = (DefaultRateClassZone) iterator.next();
			if (defaultRateClassZone.getRateClassName() != null)
			{
				if (defaultRateClassZone.getAbbreviation().equals(abbrevation) && defaultRateClassZone.getUtilityName().equals(utilitytName) && defaultRateClassZone.getRateClassName().equals(rateClassName))
				{
					isAllRateClass = true;
					break;
				}
			}
		}
		return isAllRateClass;
	}
	
	public int supplierTypeForAllRateClass(String abbrevation, String utilitytName, String rateClassName)
	{
		 boolean isAllRateClass = false;
		 int supplierType=0;
	
		for (Iterator iterator = defaultRateClassZonesList.iterator(); iterator.hasNext();)
		{
			DefaultRateClassZone defaultRateClassZone = (DefaultRateClassZone) iterator.next();
			if (defaultRateClassZone.getRateClassName() != null)
			{
				if (defaultRateClassZone.getAbbreviation().equals(abbrevation) && defaultRateClassZone.getUtilityName().equals(utilitytName) && defaultRateClassZone.getRateClassName().equals(rateClassName))
				{

					isAllRateClass = true;
					supplierType =defaultRateClassZone.getSupplierType();
					break;
				}
			}
		}
		return supplierType;
	}
}
class DefaultRateClassZone
{
	private String	utilityName		= null;
	private String	zoneName;
	private String	rateClassName;
	private String	abbreviation	= null;	//indicate state abbreviation
    private int  supplierType;
	public String getUtilityName()
	{
		return utilityName;
	}
	public void setUtilityName(String utilityName)
	{
		this.utilityName = utilityName;
	}
	public String getZoneName()
	{
		return zoneName;
	}
	public void setZoneName(String zoneName)
	{
		this.zoneName = zoneName;
	}
	public String getRateClassName()
	{
		return rateClassName;
	}
	public void setRateClassName(String rateClassName)
	{
		this.rateClassName = rateClassName;
	}
	public String getAbbreviation()
	{
		return abbreviation;
	}
	public void setAbbreviation(String abbreviation)
	{
		this.abbreviation = abbreviation;
	}
	public int getSupplierType()
	{
		return supplierType;
	}
	public void setSupplierType(int supplierType)
	{
		this.supplierType = supplierType;
	}
}
