/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 22/12/2015
 * Time: 3:07 PM
 * Version 1.0
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

namespace Dom_ClientSanityTest
{
	/// <summary>
	/// Description of SubmitComment.
	/// Scenario: Client log into Nas Domestic portal and submit comment based on the new created request number
	/// </summary>
	[TestModule("29FCB2B7-EEA6-4CE7-9EAE-6E0B06193A8A", ModuleType.UserCode, 1)]
	public class SubmitComment : ITestModule
	{
		
	#region Variables
		string _varUid = "";
		[TestVariable("B8415B09-A0E5-413F-9121-40BBEDC88E0C")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varPwd = "";
		[TestVariable("D8CA8838-7D13-48A5-85C7-17689640AA67")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("A1305469-3041-45AF-BE69-FE2263D957D8")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("F575C3F4-B817-49FB-B560-AD09865A02FE")]
		public string varUrl
		{
			get { return _varUrl; }
			set { _varUrl = value; }
		}
		
	#endregion
	
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static Dom_SanityTestRepository repo = Dom_SanityTestRepository.Instance;
		
		static SubmitRequest instance = new SubmitRequest();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public SubmitComment()
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
				
			Delay.Milliseconds(100);
			const string ClientCommentApp = "Test comment- Sent to appraiser";
			const string ClientCommentInt = "Test comment- Log Internal";
			
			//Search Nas Request and post comment to appraiser
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			Delay.Milliseconds(100);
			
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);	
			
			repo.DomNasHome.MenuDisplay.CommentOnRequest.Click();
			repo.DomNasHome.MenuDisplay.SendCommentToAppraiserFromClient.Click();
			repo.DomNasHome.MenuDisplay.CommentText.PressKeys(ClientCommentApp);
			repo.DomNasHome.MenuDisplay.SubmitComment.Click();
			Delay.Milliseconds(100);
			
			string postComment1 = repo.DomNasHome.MenuDisplay.CommentPosted.InnerText.Trim();
			
			//Report Sent to Appraiser Comment submit status
			Report.Log(ReportLevel.Success,"Validation", "Comment successfully submitted.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.CommentsWereSuccessfullySubmitted);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(ClientCommentApp, postComment1);
				
			//Post internal comment
			repo.DomNasHome.MenuDisplay.LogClientComment.Click();
			repo.DomNasHome.MenuDisplay.CommentText.PressKeys(ClientCommentInt);
			repo.DomNasHome.MenuDisplay.SubmitComment.Click();
			Delay.Milliseconds(100);
			
			string postComment2 = repo.DomNasHome.MenuDisplay.CommentPosted.InnerText.Trim();
			
			//Report Internal Comment submit status
			Report.Log(ReportLevel.Success,"Validation", "Comment successfully submitted.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.CommentsWereSuccessfullySubmitted);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(ClientCommentInt, postComment2);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
					
		}
	}
}
