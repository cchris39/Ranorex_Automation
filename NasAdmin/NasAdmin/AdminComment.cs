/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 05/02/2016
 * Time: 4:29 PM
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

namespace NasAdmin
{
	/// <summary>
	/// Description of AdminComment.
	/// </summary>
	[TestModule("E457FB70-0B82-48AC-9A11-0A0F858135F1", ModuleType.UserCode, 1)]
	public class AdminComment : ITestModule
	{
		/// <summary>
		/// Using the NasAdminRepository repository.
		/// </summary>
		public static NasAdminRepository repo = NasAdminRepository.Instance;
		
		static AdminComment instance = new AdminComment();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public AdminComment()
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
			
		string _varNasNbr = "";
		[TestVariable("8AF71B4A-5395-451E-BD19-57FF29E4E3B4")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		private const string comTypeReq = "Comment Type is required to submit a comment to the Client.";
		private const string comTypeSel = "Update";
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
			
			/*
			Host.Local.OpenBrowser(varUrl, "IE", "", false, false, false, false, false);
			Delay.Milliseconds(0);
			
			//Admin User log in
			repo.Dom_NasHome.Userid.PressKeys(varAdmin);
			repo.Dom_NasHome.Password.PressKeys(varPwd);
			repo.Dom_NasHome.RememberUserName.Click();
			repo.Dom_NasHome.LoginSubmit.Click();
			*/
			Delay.Milliseconds(300);
			
			
			//Search Nas Number
			repo.Dom_NasHome.NasAdminMenu.SearchFilter.Click();
			repo.Dom_NasHome.MenuDisplay.NasReqNumSearch.PressKeys(varNasNbr);      //varNasNbr
			repo.Dom_NasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Accepted Request
			repo.Dom_NasHome.NasAdminFunction.Chastatus.TagValue = "Accepted";
			repo.Dom_NasHome.NasAdminFunction.StatusChangeBtn.Click();
			Delay.Milliseconds(200);
			
			const string comToApp = "Test sent comment to Appraiser";
			const string comToInternal = "Log Internal Comment -Nas Admin";
			const string clientCommentTxt = "Test Send comments to Client";
			string comToClient = comTypeSel + ": " + clientCommentTxt;
			
						
			//INPUT to Appraiser COMMENT
			repo.Dom_NasHome.NasAdminFunction.CommentOnRequest.Click();
			
			repo.Dom_NasHome.MenuDisplay.Comment.SendCommentToAppraiserFromAdmin.Click();
			repo.Dom_NasHome.MenuDisplay.Comment.CommentText.PressKeys(comToApp);
			
			repo.Dom_NasHome.MenuDisplay.Comment.SubmitCommentBtn.Click();
			Delay.Milliseconds(200);
			
			string postComment1 = repo.Dom_NasHome.MenuDisplay.Comment.CommentSubmit.InnerText.Trim();
			
			//Report Sent to Appraiser Comment submit status
			Report.Log(ReportLevel.Success,"Validation", "Sent to Apprasiser Comment successfully submitted.");
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Comment.DivTagCommentsWereSuccessfullySu);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(comToApp, postComment1);
			
			//Verify select comment type message
			repo.Dom_NasHome.MenuDisplay.Comment.SendCommentToClientFromAdmin.Click();
			repo.Dom_NasHome.MenuDisplay.Comment.SubmitCommentBtn.Click();
						
			string reqCommType = repo.MessageFromWebpage.Text65535_CommentTypeRequired.Caption.Trim();
			
			if (reqCommType == comTypeReq)
			{
				Report.Log(ReportLevel.Success, "Success", "Comment Type required message verified.");
				repo.MessageFromWebpage.ButtonOK.Click();
			}else
			{   Report.Log(ReportLevel.Warn, "Warning", "Comment Type required message verification failed, please checked.");}
			
			
			//Input to Client Comment
			repo.Dom_NasHome.MenuDisplay.Comment.AdminCommentType.TagValue = comTypeSel;	
			repo.Dom_NasHome.MenuDisplay.Comment.AdminCommentType.Click();
			repo.Dom_NasHome.MenuDisplay.Comment.CommentText.Click();
			repo.Dom_NasHome.MenuDisplay.Comment.CommentText.PressKeys(clientCommentTxt);
			
			repo.Dom_NasHome.MenuDisplay.Comment.SubmitCommentBtn.Click();
			Delay.Milliseconds(200);
			
			TdTag comTxt2 = "/dom[@domain='test.nationwideappraisals.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//tr[3]//td[3]";
			string postComment2 = comTxt2.InnerText.Trim();
			
			//Report Sent to Appraiser Comment submit status
			Report.Log(ReportLevel.Success,"Validation", "Sent to Client Comment successfully submitted.");
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Comment.DivTagCommentsWereSuccessfullySu);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(comToClient, postComment2);
			
			//Input Log Internal Comment
			repo.Dom_NasHome.MenuDisplay.Comment.LogInternalComment.Click();
			repo.Dom_NasHome.MenuDisplay.Comment.CommentText.PressKeys(comToInternal);
			
			repo.Dom_NasHome.MenuDisplay.Comment.SubmitCommentBtn.Click();
			Delay.Milliseconds(200);
			
			TdTag comTxt3 = "/dom[@domain='uattest.nas.com']//frameset[#'mainFrameset']/frame[@name='menu_display']//tr[4]//td[3]";
			string postComment3 = comTxt3.InnerText.Trim();
			
			//Report Sent to Appraiser Comment submit status
			Report.Log(ReportLevel.Success,"Validation", "Internal Comment successfully submitted.");
			Validate.Exists(repo.Dom_NasHome.MenuDisplay.Comment.DivTagCommentsWereSuccessfullySu);
			Report.Log(ReportLevel.Success, "Validation", "Comment text match");
			Validate.AreEqual(comToInternal, postComment3);
			
			Delay.Milliseconds(200);
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			//Delay.Milliseconds(400);
			
		}
	}
}
