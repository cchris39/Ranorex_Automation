/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 04/02/2016
 * Time: 3:15 PM
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



namespace NasAdmin
{
	/// <summary>
	/// Description of AdminSearchRequest.
	/// Scenario: Admin user performed  search function;
	/// </summary>
	[TestModule("92DD10D5-2C66-4AF1-92EC-06B8FD13A5CF", ModuleType.UserCode, 1)]
	public class AdminSearchRequest : ITestModule
	{
		
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static AdminSearchRequest instance = new AdminSearchRequest();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AdminSearchRequest()
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
		
		string _varPostalCode = "";
		[TestVariable("16C7A63F-ABF5-4B58-A8AA-96EBB1F7CC3A")]
		public string varPostalCode
		{
			get { return _varPostalCode; }
			set { _varPostalCode = value; }
		}
		
		string _varProvince = "";
		[TestVariable("CD6EC14F-8F98-49A1-9B47-82834648EB84")]
		public string varProvince
		{
			get { return _varProvince; }
			set { _varProvince = value; }
		}
		
		string _varCity = "";
		[TestVariable("D54CD9E6-7E47-4EC0-86DD-A03B8E4C0BD0")]
		public string varCity
		{
			get { return _varCity; }
			set { _varCity = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("B959D310-A651-40EE-88C4-C7AA7627B213")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varCreationDate = "";
		[TestVariable("10068BD5-63AA-49B7-9858-94A26CD7DAAF")]
		public string varCreationDate
		{
			get { return _varCreationDate; }
			set { _varCreationDate = value; }
		}
		
		string _varRefNbr = "";
		[TestVariable("15AA0CAC-7668-4180-A089-23C3B86C669B")]
		public string varRefNbr
		{
			get { return _varRefNbr; }
			set { _varRefNbr = value; }
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
						
			//repo.NotForThisSite.Click();
			
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			Delay.Milliseconds(100);	
			
			//Input Nas Number to search
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Report Search by Nas Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr + " was found by searching Nas number.");
			///Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr +  );
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchNbr = repo.Dom_NasHome.MenuDisplay.DivTagNasNbr.InnerText.Trim();
			
			Report.Log(ReportLevel.Info, "Validation", "Nas number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchNbr, varNasNbr);
				
			//Search By Client Reference Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.ClientRefNum.PressKeys(varRefNbr);   //varRefNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Report Search by Client Referench Number status
			Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr  + " was found by searching client Reference number: " + varRefNbr);
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchRefNbr = repo.Dom_NasHome.MenuDisplay.DivTagNasNbr.InnerText.Trim();
			
			//Report.Log(ReportLevel.Info, "Validation", "Nas Number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchRefNbr, varNasNbr);
			Delay.Milliseconds(200);
			
			//Search by Date
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.CreateDateFrom.Element.SetAttributeValue("TagValue", varCreationDate); //varCreationDate
			repo.Dom_NasHome.MenuDisplay.CreateDateTo.Element.SetAttributeValue("TagValue",varCreationDate);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			string popOpMsg = repo.MessageFromWebpage.Text65535.Caption.Trim();
						
			//Validate pop op message
			if (repo.MessageFromWebpage.Text65535 != null)
			{
				Report.Log(ReportLevel.Info, "Validation", "Search Pop Op message verified.");
				Report.Info("Pop op Message is: " + popOpMsg);
				repo.MessageFromWebpage.ButtonOK.Click();
			}else
			{
				Report.Log(ReportLevel.Info, "Validation", "No Pop Op message presented.");
			}
			
			
			//Input the missing city Information to search by Date range
			repo.Dom_NasHome.MenuDisplay.CountryCode.Click();
			repo.Dom_NasHome.MenuDisplay.CountryCode.TagValue = "CAN";
			Delay.Milliseconds(400);
			//repo.Dom_NasHome.MenuDisplay.Province.Click();
			repo.Dom_NasHome.MenuDisplay.Province.TagValue = varProvince;      //varProvince
			Delay.Milliseconds(400);
			repo.Dom_NasHome.MenuDisplay.City.PressKeys(varCity);        //varCity
			Delay.Milliseconds(300);
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Loop the search result table to validate request found
			for (int i = 1; i <= 11; i++)
				{
					string varTRrow = i.ToString();
					string XpathNasNbrFound = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']/body/form[@name='f_form']/?/?/table[@class='changeOnHover']//tr[#'trRow" + varTRrow + "']//td[2]//div";
					
					
					
					Ranorex.DivTag nasNbr_Div = XpathNasNbrFound;
			
					string searchDateNbr = nasNbr_Div.InnerText.Trim();
					//MessageBox.Show(searchDateNbr);
					if (searchDateNbr == varNasNbr)
					{
						Report.Log(ReportLevel.Success, "Validation", "Request Number: " + varNasNbr  + " was found by searching requested date: " + varCreationDate); 	  //varNasNbr
						Validate.AreEqual(searchDateNbr, varNasNbr);
						break;
					}
				}
				
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(300);
			
			} // End of Void Main
		
	} //End of Class
	
	

} //End Of NameSpace
		