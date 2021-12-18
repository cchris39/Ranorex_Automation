/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 22/01/2016
 * Time: 3:35 PM
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

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserSearchRequest.
	/// </summary>
	[TestModule("C502A24A-519E-4872-BB59-7882A80B1CEF", ModuleType.UserCode, 1)]
	public class AppraiserSearchRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserSearchRequest instance = new AppraiserSearchRequest();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserSearchRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}

		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		#region Varilables
		
		string _varAppraiser = "";
		[TestVariable("AA9A04F4-065E-44E2-9C6A-5487AAD8B6F7")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("BC42EDD9-1907-42C5-ADE0-7EC36F2409DD")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("E6DFACA1-D3FA-4F8B-B5C3-452391D292DD")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("8808E13D-C395-48D5-B644-7275DC61C684")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("DE2570FB-A399-400F-8C7E-1E6C1653C6F2")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
		#endregion
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
			
			//Host.Local.ClearBrowserCookies("IE");
			//Delay.Milliseconds(100);
			
			//Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
			
			//Appraiser log in
			//repo.DomNasHome.Userid.PressKeys(varAppraiser);
			//repo.DomNasHome.Password.PressKeys(varPwd);
			//repo.DomNasHome.Submit.Click();
			//Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Report Search by Nas Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr + " was found by searching Nas number.");
			///Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr +  );
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchNbr = repo.DomNasHome.MenuDisplay.LabelTagNasNum.InnerText.Trim();
			
			Report.Log(ReportLevel.Info, "Validation", "Nas number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchNbr, varNasNbr);
				
			//Search By Client Reference Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.ClientRefNum.PressKeys(varRefNbr);   //varRefNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Report Search by Client Referench Number status
			Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr  + " was found by searching client Reference number: " + varRefNbr);
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchRefNbr = repo.DomNasHome.MenuDisplay.LabelTagNasNum.InnerText.Trim();
			
			Report.Log(ReportLevel.Info, "Validation", "Nas Number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchRefNbr, varNasNbr);
			
			//Search by Date
			var curDate = System.DateTime.Today.ToString("yyyy-MM-dd");
			var preDate = System.DateTime.Today.AddDays(-10).ToString("yyyy-MM-dd");
				
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.CreateDateFrom.Element.SetAttributeValue("TagValue", preDate); 
			repo.DomNasHome.MenuDisplay.CreateDateTo.Element.SetAttributeValue("TagValue",curDate);
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Loop the search result table to validate request found | UAT Part
			Ranorex.TableTag searchResult_table = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//div[#'subContainer']/table[2]";
			
			if (searchResult_table == null)
			{
				//Report failure searching by date after loop over the result table
				Report.Log(ReportLevel.Failure, "Validation", "Request Number: " + varNasNbr  + " was not found by searing date range: " + preDate + " To " + curDate);
				Delay.Milliseconds(100);	
			}
			
			for (int i = 1; i <= 11; i++)
				{
					string varTRrow = i.ToString();
				//UAT Part
				//string XpathNasNbrFound = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//div[#'subContainer']/table[2]//" + "tr[" + varTRrow + "]//td[2]//label[]";       
				string XpathNasNbrFound = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//tr[#'trRow$']//td[2]//label[]".Replace("$",varTRrow).Trim();                                        //Path changed (03.23.2016)
					
										
					Ranorex.LabelTag nasNbr_Label = XpathNasNbrFound;
			
					string searchDateNbr = nasNbr_Label.InnerText.Trim();
					//MessageBox.Show(searchDateNbr);
					
					if (searchDateNbr == varNasNbr) {
						Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr  + " was found by searing date range: " + preDate + " To " + curDate); 	  //varNasNbr
						Validate.AreEqual(searchDateNbr, varNasNbr);
						break;
					} 
				}
			
				
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(200);
			
		}
	}
}

