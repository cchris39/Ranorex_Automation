/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 23/12/2015; 04/01/2016
 * Time: 3:26 PM
 * Version: 1.0; 1.1 | 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Data;
using System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of SearchRequest.
	/// Scenario: Search submitted request by different searching methods
	/// </summary>
	[TestModule("D76D8F8B-28C0-440E-8E60-DA65DB076525", ModuleType.UserCode, 1)]
	public class SearchRequest : ITestModule
	{
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
		static SearchRequest instance = new SearchRequest();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SearchRequest()
		{
			// Do not delete - a parameterless constructor is required!
		}
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
		
	#region Variables
		
		string _varUid = "";
		[TestVariable("1C26A9CA-02AE-49AA-BFCE-2659A4709CEF")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
	
		string _varPwd = "";
		[TestVariable("89401E95-3A7F-4A60-B872-93001151B37E")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("729B2BE3-5BAD-4145-8B26-7E104AA5CD34")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("18AF2D38-745A-48C4-965D-6757CB01967D")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("34300E08-2BD2-4634-A532-9B6E2844DC7D")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _varUrl = "";
		[TestVariable("499B01B7-ED51-495E-8E9F-D6036A0B2089")]
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
			
			/*/
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(200);
			
			//Login
			repo.DomNasHome.Userid.PressKeys(varUid);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			/*/
			
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
			
			//Report.Log(ReportLevel.Info, "Validation", "Nas Number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchRefNbr, varNasNbr);
			
			//Search by Date
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.CreateDateFrom.Element.SetAttributeValue("TagValue", varCreationDate); //varCreationDate
			repo.DomNasHome.MenuDisplay.CreateDateTo.Element.SetAttributeValue("TagValue",varCreationDate);
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(300);
			
			//Loop the search result table to validate request found
			for (int i = 1; i <= 11; i++)
				{
					string varTRrow = "#'trRow" + i.ToString() + "'";
					//This run at UAT
					string XpathNasNbrFound = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//" + "tr[" + varTRrow + "]//td[2]//label[]";
					
					Ranorex.LabelTag nasNbr_Label = XpathNasNbrFound;
			
					string searchDateNbr = nasNbr_Label.InnerText.Trim();
					
					
					if (searchDateNbr == varNasNbr) {
						Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr  + " was found by searching requested date: " + varCreationDate); 	  //varNasNbr
						Validate.AreEqual(searchDateNbr, varNasNbr);
						break;
					} 
				}
				
			//Report failure searching by date after loop over the result table
			//Report.Log(ReportLevel.Failure, "Validation", "Request Number: " + varNasNbr  + " was not found by searing request date.");
			//Delay.Milliseconds(100);	
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
			
		}
	}
}
