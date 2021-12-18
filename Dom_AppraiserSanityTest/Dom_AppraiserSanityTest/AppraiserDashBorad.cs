/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 18/01/2016
 * Time: 12:03 PM
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
using System.Data;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Dom_AppraiserSanityTest
{
	/// <summary>
	/// Description of AppraiserDashBorad.
	/// Scenario: Appraiser dashboard function: view page and comment link checked; Qucik search function; 
	/// //		  Get total number for active requests status and verified the status link and results show in the table.
	/// </summary>
	[TestModule("5D374C51-DAF0-4982-A8C9-C691C001D9B2", ModuleType.UserCode, 1)]
	public class AppraiserDashBorad : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserDashBorad instance = new AppraiserDashBorad();
		
		public DataTable activeTable = new DataTable();
		
	#region Variables
	
	string _varNasNbr = "";
	[TestVariable("1DB27AD6-B21D-48D5-A548-6081F975B177")]
	public string varNasNbr
	{
		get { return _varNasNbr; }
		set { _varNasNbr = value; }
	}
	
	string _varUrl = "";
	[TestVariable("7A9F019C-6890-432E-9699-6FCBF4B5942C")]
	public string varUrl
	{
		get { return _varUrl; }
		set { _varUrl = value; }
	}
	
	string _varAppraiser = "";
	[TestVariable("BC6B8250-8E98-4E23-9736-A629CDAF0E5B")]
	public string varAppraiser
	{
		get { return _varAppraiser; }
		set { _varAppraiser = value; }
	}
	
	string _varPwd = "";
	[TestVariable("414AF298-B739-4125-ABF4-4126CAC309D1")]
	public string varPwd
	{
		get { return _varPwd; }
		set { _varPwd = value; }
	}
	
	string _varPropertyAddress = "";
	[TestVariable("08D22BB6-0392-44C0-BE82-7955A472DCAC")]
	public string varPropertyAddress
	{
		get { return _varPropertyAddress; }
		set { _varPropertyAddress = value; }
	}
	
	string _varRefNbr = "";
	[TestVariable("4C5777EF-B887-4106-AE8D-4800637BD4DE")]
	public string varRefNbr
	{
		get { return _varRefNbr; }
		set { _varRefNbr = value; }
	}
	
	#endregion

		
		public void KillBrowser(string IE)
		{
		}

		public void ClearBrowserCookies(string IE)
		{
		}
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserDashBorad()
		{
			// Do not delete - a parameterless constructor is required!
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
			
			Host.Local.ClearBrowserCookies("IE");
			Delay.Milliseconds(100);
			
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(300);
			
			//Appraiser log in
			repo.DomNasHome.Userid.PressKeys(varAppraiser);
			repo.DomNasHome.Password.PressKeys(varPwd);
			repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(300);
			
			//Check if the E&O EXPIRY Warning exist 
						
			//Search From DahBoard Search Box by searching Nas Number
			repo.DomNasHome.MenuDisplay.AppraiserDashBoardQuicksearch.PressKeys(varNasNbr);   //varNasNbr
			repo.DomNasHome.MenuDisplay.SearchGoBtn.Click();
			
			//Report Search by Nas Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr + " was found by searching Nas number.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchNbr = repo.DomNasHome.MenuDisplay.LabelTagNasNum.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Validation", "Nas number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchNbr, varNasNbr);
			
			//Search From DahBoard Search Box by searching clientRefence Number
			repo.DomNasHome.Dashboard.Click();
			Delay.Milliseconds(300);
			
			repo.DomNasHome.MenuDisplay.AppraiserDashBoardQuicksearch.PressKeys(varRefNbr);
			repo.DomNasHome.MenuDisplay.SearchGoBtn.Click();
			
			//Report Search by Client Referench Number status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr  + " was found by searching client Reference number: " + varRefNbr);
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchRefNbr = repo.DomNasHome.MenuDisplay.LabelTagNasNum.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Validation", "Nas Number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchRefNbr, varNasNbr);
			
			//Search From DahBoard Search Box by searching address
			repo.DomNasHome.Dashboard.Click();
			Delay.Milliseconds(300);
			
			repo.DomNasHome.MenuDisplay.AppraiserDashBoardQuicksearch.PressKeys(varPropertyAddress);
			repo.DomNasHome.MenuDisplay.SearchGoBtn.Click();
			
			//Report Search by Property Address status
			Report.Log(ReportLevel.Info, "Validation", "Request Number: " + varNasNbr  + " was found by searching property address " + varPropertyAddress);
			Validate.Exists(repo.DomNasHome.MenuDisplay.StrongTag1RecordSFound);	
			
			string SearchAddressNbr = repo.DomNasHome.MenuDisplay.LabelTagNasNum.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Validation", "Nas Number: " + varNasNbr  + " is match");
			Validate.AreEqual(SearchAddressNbr, varNasNbr);
			
			//Sample active Request Search  (Sample)
			repo.DomNasHome.Dashboard.Click();
			Delay.Milliseconds(200);
			
			repo.DomNasHome.MenuDisplay.ActiveRequestSearch.PressKeys(varNasNbr);    //varNasNbr
			Keyboard.Press("{Return}");
			Delay.Milliseconds(100);
						
			string sampleNasNbr = repo.DomNasHome.MenuDisplay.SampleRequestNbr.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Validation", "Sample reuqest search works properly. Nas number is match");
			Validate.AreEqual(sampleNasNbr, varNasNbr);
			
			//Get total active request for each status
			string [] pieStatus = new string [5] {"New", "Accepted", "Appointment", "Under Review", "Pending"};
						
			string [] statusTotal = new string [5];
			statusTotal[0] = repo.DomNasHome.MenuDisplay.ATagNew.InnerText.Replace(" New", "").Trim();
			statusTotal[1] = repo.DomNasHome.MenuDisplay.ATagAccepted.InnerText.Replace(" Accepted", "").Trim();
			statusTotal[2] = repo.DomNasHome.MenuDisplay.ATagAppointment.InnerText.Replace(" Appointment", "").Trim();
			statusTotal[3] = repo.DomNasHome.MenuDisplay.ATagUnderReview.InnerText.Replace(" Under Review","").Trim();
			statusTotal[4] = repo.DomNasHome.MenuDisplay.ATagPending.InnerText.Replace(" Pending", "").Trim();
			
			
			string[] actStatus = new string [5];
			
			for (int i = 0; i <= 4; i++)
       		{
			 	
				//Loop through different active status and check the if the link work properly
			 							
			 					switch (i) 
			 					{
									case 0:
												repo.DomNasHome.MenuDisplay.ATagNew.Click();
												Delay.Milliseconds(200);
												
												if (statusTotal[i] == "0")
												{
													actStatus[i] = "No results found";
													break;
												}
													actStatus[i] = repo.DomNasHome.MenuDisplay.ActiveRequestStatus.InnerText.Trim();
													break;
													
												
									case 1:			
												repo.DomNasHome.MenuDisplay.ATagAccepted.Click();
												Delay.Milliseconds(200);
												
												if (statusTotal[i] == "0")
												{
													actStatus[i] = "No results found";
													break;
												}
													actStatus[i] = repo.DomNasHome.MenuDisplay.ActiveRequestStatus.InnerText.Trim();
													break;
													
									
									case 2:			
												repo.DomNasHome.MenuDisplay.ATagAppointment.Click();
												Delay.Milliseconds(300);
												
												if (statusTotal[i] == "0")
												{
													actStatus[i] = "No results found";
													break;
												}
													actStatus[i] = repo.DomNasHome.MenuDisplay.ActiveRequestStatus.InnerText.Trim();
													break;
													
													
									case 3:			
												repo.DomNasHome.MenuDisplay.ATagUnderReview.Click();
												Delay.Milliseconds(200);
												
												if (statusTotal[i] == "0")
												{
													actStatus[i] = "No results found";
													break;
												}
													
													actStatus[i] = repo.DomNasHome.MenuDisplay.ActiveRequestStatus.InnerText.Trim();
													break;
													
									
									case 4:			
												repo.DomNasHome.MenuDisplay.ATagPending.Click();
												Delay.Milliseconds(200);
												
												if (statusTotal[i] == "0")
												{
													actStatus[i] = "No results found";
													break;
												}
													actStatus[i] = repo.DomNasHome.MenuDisplay.ActiveRequestStatus.InnerText.Trim();
													break;
													
																					
									default: 
			 									Report.Log(ReportLevel.Info, "Validation", "No active status link being checked." );
												Delay.Milliseconds(100);
			 									break;
								}//End of switch case												
												
												//Validation and report result
												if (actStatus[i] == "No results found")
												{
													//var noFound = repo.DomNasHome.MenuDisplay.NoResultsFoundDashBoard.InnerText.Trim();
													Report.Log(ReportLevel.Success,"Validation", "Active request " + pieStatus[i] + " link checked: Pass");
													Report.Log(ReportLevel.Info, "Validation", statusTotal[i] + " active requests for: " + pieStatus[i]);
													Validate.Exists(repo.DomNasHome.MenuDisplay.NoResultsFoundDashBoard);
													Delay.Milliseconds(100);
												}else
												{
													Report.Log(ReportLevel.Info, "Validation", "Active request for " + pieStatus[i]+ " status: " + statusTotal[i]);
													Report.Log(ReportLevel.Success,"Validation", "Active request " + pieStatus[i] + " link checked: Pass");
													Validate.AreEqual(actStatus[i], pieStatus[i]);
													Delay.Milliseconds(100);
												}
										 						
			 		}	//End of for loop
			 	
					//Check Avtive request comment link
					repo.DomNasHome.MenuDisplay.ViewAllActiveRequestsBtn.Click();
					Delay.Milliseconds(300);
					
					repo.DomNasHome.MenuDisplay.ActiveCommentBtn.Click();
					Delay.Milliseconds(100);
					Report.Log(ReportLevel.Success, "Validation", "Active request comment link checked: " + "Pass");
					Report.Log(ReportLevel.Info, "Validation", "Comment page verified");
					Validate.Exists(repo.DomNasHome.MenuDisplay.AppraiserCommentTableTag);	
					
					//Check Avtive request Nas Number link
					repo.DomNasHome.Dashboard.Click();
					Delay.Milliseconds(300);
					
					repo.DomNasHome.MenuDisplay.ActiveViewTagNasNbr.Click();
					Delay.Milliseconds(100);
					Report.Log(ReportLevel.Success, "Validation", "Active request view request link checked: " + "Pass");
					Report.Log(ReportLevel.Info, "Validation", "Search result page verified");
					Validate.Exists(repo.DomNasHome.MenuDisplay.H1TagAppraisalSearchRes);	
					
					//Close Browser
					//Host.Local.KillBrowser("IE");
					//Delay.Milliseconds(300);
			
		}
		
		
		
	}
}
			