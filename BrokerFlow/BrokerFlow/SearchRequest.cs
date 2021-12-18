/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 4:15 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace BrokerFlow
{
	/// <summary>
	/// Description of SearchRequest.
	/// </summary>
	[TestModule("9F48911F-017C-4976-A7F4-7DFE71DFA0BB", ModuleType.UserCode, 1)]
	public class SearchRequest : ITestModule
	{
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("C19D0853-540F-4D70-B573-1F7BED98EA02")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		
		string _varRefNbr = "";
		[TestVariable("E7708562-1CCE-4F0F-A377-C9CA59780CBA")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
		}
		
		
		string _varCreationDate = "";
		[TestVariable("3D74F521-223A-494D-B23C-D7A9D3DBF4FB")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		#endregion
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
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
					
					//Remember to chnage the domin name if run in production !!!
					//string XpathNasNbrFound = "/dom[@domain='www.nationwideappraisals.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//" + "tr[" + varTRrow + "]//td[2]//label[]";
					
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
