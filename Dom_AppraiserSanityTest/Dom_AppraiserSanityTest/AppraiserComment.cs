/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 25/01/2016
 * Time: 12:00 PM
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
	/// Description of AppraiserComment.
	/// Scenario: Appraiser submit two comment: 1. Need additional conatct/address; 2. Left message for contact
	/// </summary>
	[TestModule("9B1961B3-F146-434E-8E8A-54999EC66997", ModuleType.UserCode, 1)]
	public class AppraiserComment : ITestModule
	{
		/// <summary>
		/// Using the Dom_AppraiserSanityTestRepository repository.
		/// </summary>
		public static Dom_AppraiserSanityTestRepository repo = Dom_AppraiserSanityTestRepository.Instance;
		
		static AppraiserComment instance = new AppraiserComment();
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AppraiserComment()
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
		[TestVariable("DCDD9C23-CA81-4DE2-A981-90FF032781C9")]
		public string varAppraiser
		{
			get { return _varAppraiser; }
			set { _varAppraiser = value; }
		}
		
		string _varPwd = "";
		[TestVariable("E0A70A5E-128B-4804-8662-68C4D0771EA8")]
		public string varPwd
		{
			get { return _varPwd; }
			set { _varPwd = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("241BB6E7-90CF-4979-A572-5AF39BBFA19C")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		string _varUrl = "";
		[TestVariable("FA15ECDB-9C2C-45F2-B6C4-D26F4266D62F")]
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
			//Delay.Milliseconds(100);

			//WebDocument DomWebDocu = "/dom[@domain='uattest.nas.com']";

			//Appraiser log in
			//repo.DomNasHome.Userid.PressKeys(varAppraiser);
			//repo.DomNasHome.Password.PressKeys(varPwd);
			//repo.DomNasHome.Submit.Click();
			Delay.Milliseconds(100);
			
			//Search By Nas Number
			repo.DomNasHome.SearchFilter.Click();
			repo.DomNasHome.MenuDisplay.ViewUserReq.Click();
			repo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(varNasNbr);     // varNasNbr
			repo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(100);
			
			//Accepted Request
			string chgStatus = "Accepted";
			
			repo.DomNasHome.MenuDisplay.Chastatus.Element.SetAttributeValue("TagValue", chgStatus);
			repo.DomNasHome.MenuDisplay.ChgStatusBtn.Click();
			Delay.Milliseconds(100);
			
			Report.Log(ReportLevel.Info, "Validation", "Status changed for: " + varNasNbr);
			Validate.Exists(repo.DomNasHome.MenuDisplay.StatusChangedFromNew);
			
			//Appraiser submit comments
			repo.DomNasHome.MenuDisplay.CommentOnRequest.Click();
			Delay.Milliseconds(100);
			
			const string commMoreContact = "Need additional contact/address information";
			const string commMsgContact = "Left message for the contact(s)";
			
			//Submit comment option 1
			repo.DomNasHome.AppraiserCommentsOptions.RequireVerification.Click();
			repo.DomNasHome.MenuDisplay.CommentText.PressKeys(commMoreContact);
			repo.DomNasHome.MenuDisplay.SubmitComment.Click();
			
			string postComm1 = repo.DomNasHome.MenuDisplay.CommentPosted.InnerText.Trim();
				
			Report.Log(ReportLevel.Success,"Validation", "Comment successfully submitted.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.CommentsWereSuccessfullySubmitted);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(commMoreContact, postComm1);
			
			//Submit comment option 2
			repo.DomNasHome.AppraiserCommentsOptions.MessageLeft.Click();
			repo.DomNasHome.MenuDisplay.SubmitComment.Click();
			
			string postComm2 = repo.DomNasHome.MenuDisplay.CommentPosted.InnerText.Trim();
				
			Report.Log(ReportLevel.Success,"Validation", "Comment successfully submitted.");
			Validate.Exists(repo.DomNasHome.MenuDisplay.CommentsWereSuccessfullySubmitted);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(commMsgContact, postComm2);
				
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(200);
			
			
		}
	}
}
