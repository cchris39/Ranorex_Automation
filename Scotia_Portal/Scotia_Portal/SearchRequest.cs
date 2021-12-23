/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-10
 * Time: 4:27 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Scotia_Portal
{
	/// <summary>
	/// Description of SearchRequest.
	/// </summary>
	[TestModule("AD79676D-B4D1-4036-B5D5-DA5A6A0107DF", ModuleType.UserCode, 1)]
	public class SearchRequest : ITestModule
	{
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
			static SearchRequest instance = new SearchRequest();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SearchRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("65E7336B-1A61-4C8A-ADBC-F667E8475132")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varStatus = "";
		[TestVariable("86089769-D13B-441F-9A87-937CAD7A1B2D")]
		public string varStatus
		{
			get { return _varStatus; }
			set { _varStatus = value; }
		}
		
		string _varDate = "";
		[TestVariable("544F37F5-B707-4C0B-998D-707136D5FE2B")]
		public string varDate
		{
			get { return _varDate; }
			set { _varDate = value; }
		}
		
		string _varPostal = "";
		[TestVariable("5F419143-E985-45CC-B154-4DB56AABA088")]
		public string varPostal
		{
			get { return _varPostal; }
			set { _varPostal = value; }
		}
		
		string _varAppName = "";
		[TestVariable("A384BD9E-2CDD-453D-BD0E-82A1FC1283F3")]
		public string varAppName
		{
			get { return _varAppName; }
			set { _varAppName = value; }
		}
		
		string _varAddress = "";
		[TestVariable("B2B758EB-091D-4FCE-98BB-C6E287C52A2F")]
		public string varAddress
		{
			get { return _varAddress; }
			set { _varAddress = value; }
		}
		
		string _varRefNumber = "";
		[TestVariable("C6EFD412-A153-42A6-AD18-10E2A4C034F5")]
		public string varRefNumber
		{
			get { return _varRefNumber; }
			set { _varRefNumber = value; }
		}
		
		string _varCity = "";
		[TestVariable("1BCE2E44-AAAF-407B-8329-2075B9252198")]
		public string varCity
		{
			get { return _varCity; }
			set { _varCity = value; }
		}
		
		string _varProv = "";
		[TestVariable("24BE42F2-2467-4E90-86A2-57D518C86478")]
		public string varProv
		{
			get { return _varProv; }
			set { _varProv = value; }
		}
		
		#endregion
		
    	public void search()
    	{
    		repo.DomScotia.MainMenu.SearchFilter.Click();
			Delay.Milliseconds(100);
		}
    	
    	public void findByNasNbr(string nbr)
		{
			
			//Search BY Nas Number
			repo.DomScotia.SearchFilter.NasReqNum.PressKeys(nbr);
			repo.DomScotia.SearchFilter.ViewGroupReq.Click();
			repo.DomScotia.SearchFilter.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			string nasNbrFound = repo.DomScotia.SearchResults.NasNumberInTBL.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Success", "Request " + nbr + " found.");
			Validate.AreEqual(nasNbrFound, nbr);
			
		}
    	
    	public void findByDate(string dateFrom, string dateTo)
		{
			//Search BY Date Range
			Delay.Milliseconds(100);
			repo.DomScotia.SearchFilter.CreateDateFrom.Click();
			repo.DomScotia.SearchFilter.CreateDateFrom.TagValue = dateFrom;
			repo.DomScotia.SearchFilter.CreateDateTo.Click();
			repo.DomScotia.SearchFilter.CreateDateTo.TagValue = dateTo;
			repo.DomScotia.SearchFilter.ViewGroupReq.Click();
			repo.DomScotia.SearchFilter.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
		}
    	
    	public void findByRefNbr(string refNbr)
    	{
    		repo.DomScotia.MainMenu.SearchFilter.Click();
			Delay.Milliseconds(100);
			repo.DomScotia.SearchFilter.ClientRefNum.PressKeys(refNbr);
			
    	}
    	
    	public void findByAddress(string prov, string city , string address)
    	{
    		repo.DomScotia.SearchFilter.Province.TagValue = prov;
    		repo.DomScotia.SearchFilter.City.PressKeys(city);
    		repo.DomScotia.SearchFilter.Address.PressKeys(address);
    		
			Delay.Milliseconds(100);
    	}
    	
    	public void findByPoCode(string poCode)
    	{
    		repo.DomScotia.SearchFilter.Postal.PressKeys(poCode);
			Delay.Milliseconds(100);
    	}
    	
    	public void findByAppName(string appName)
    	{
    		
    		repo.DomScotia.SearchFilter.ClientName.PressKeys(appName);
    		Delay.Milliseconds(100);
    		
    	}
    	public bool foundNasNbr(string nbr)
    	{
    		int flag = 0;
    		for (int i=1; i<=10; i++)
    		{
    			Ranorex.LabelTag foundResult = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//div[#'subContainer']/table[@class='changeOnHover']//tr['trRow$']//td[2]//label[]".Replace("$", i.ToString());
    			string nasNbrFound = foundResult.InnerText.Trim();
    			if(nasNbrFound == nbr)
    			{ 	
    				flag = 1;
    				break;
    				
    			}
    		}
    		if (flag == 1)
    		{ return true;}
    		else {return false;}
    		
    	}
    	
    	
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//GET CURRENT DATE VALUE
			var curDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			Delay.Milliseconds(100);
			
			
			//Search By Nas number
			search();
			findByNasNbr(varNasNbr);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by Nas number: " + varNasNbr);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			//Serch By Application Number
			search();
			findByRefNbr(varRefNumber);
			findByDate(varDate, curDate);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by Application number: " + varRefNumber);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			//Search By Address
			search();
			findByAddress(varProv, varCity, varAddress);
			findByDate(varDate, curDate);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by property address: " + varAddress);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			//Search By Postal Code
			search();
			findByPoCode(varPostal);
			findByDate(varDate, curDate);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by property postal code: " + varPostal);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			//Search By Applicant Name
			search();
			findByAppName(varAppName);
			findByDate(varDate, curDate);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by property postal code: " + varAppName);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			//Search By Date
			search();
			findByDate(varDate, curDate);
			
			if(foundNasNbr(varNasNbr))
			{
				Report.Log(ReportLevel.Info, "Searching request by creation date: " + varDate);
				Report.Log(ReportLevel.Success, "Success", "Request " + varNasNbr + " found.");
			}else
			{ Report.Log(ReportLevel.Failure, "Fail", varNasNbr + " search failed, please check.");}
			Delay.Milliseconds(100);
			
			
			//Close Browser
				//Host.Local.KillBrowser("IE");
				Delay.Milliseconds(200);
			
		}
	}
}
