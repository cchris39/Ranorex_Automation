/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 14/03/2016
 * Time: 4:21 PM
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
	/// Description of SubmitComment.
	/// </summary>
	[TestModule("42E61BA2-6341-4663-A4A1-123014D528D7", ModuleType.UserCode, 1)]
	public class SubmitComment : ITestModule
	{
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("49820A65-D4A3-42F4-80C0-E4A11D80D608")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Using the Dom_SanityTestRepository repository.
		/// </summary>
		public static BrokerFlowRepository repo = BrokerFlowRepository.Instance;
		
		static SubmitComment instance = new SubmitComment();
		
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
			
			Delay.Milliseconds(200);
		}
	}
}
