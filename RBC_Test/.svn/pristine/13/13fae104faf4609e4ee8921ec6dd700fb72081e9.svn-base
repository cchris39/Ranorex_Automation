/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-04-29
 * Time: 9:23 AM
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

namespace RBC_Test
{
	/// <summary>
	/// Description of Comment.
	/// </summary>
	[TestModule("7EF329DA-AE21-463C-A5D4-8C4629A3C43C", ModuleType.UserCode, 1)]
	public class Comment : ITestModule
	{
		
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static Comment instance = new Comment();
		
		#region Varilables
		
		string _varNasNbr = "";
		[TestVariable("B809FBC4-AB59-4141-8744-AF1FBCCA6D56")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		private const string commentTxt = "Test comment enter here, Thank you.";
		private string commentID,comment, commentDate;
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Comment()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		public void GeneralComment()
		{
			repo.RBC_Portal.Main.CommentValueAppeal.Click();
			Delay.Milliseconds(100);
			
			//Post Comment to Appraiser and Nas
			repo.RBC_Portal.Main.SendCommentToAppraiserNas.Click();
			repo.RBC_Portal.Main.CommentText.PressKeys(commentTxt);
			repo.RBC_Portal.Main.CommentSubmit.Click();
			Delay.Milliseconds(300);
			
			commentID = repo.RBC_Portal.Main.CommentIDTag.InnerText.Trim();
			comment = repo.RBC_Portal.Main.CommentTextTag.InnerText.Trim();
			commentDate = repo.RBC_Portal.Main.CommentPostDate.InnerText.Trim();
			
			//Report Comment submit status
			Report.Log(ReportLevel.Success, "Validation", "Comment was successfully submitted." );
			Validate.Exists(repo.RBC_Portal.Main.CommentsSuccessfulSubmit);
			Report.Log(ReportLevel.Info, "Validation", "Comment text match.");
			Validate.AreEqual(comment, commentTxt);
			Report.Log(ReportLevel.Info, "Information", "Comment ID: " + commentID);
			Report.Log(ReportLevel.Info, "Information", "Comment posted on: " + commentDate);
			
				
		}

		
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//Search BY Nas Number
			repo.RBC_Portal.SearchFilter.Click();
			Delay.Milliseconds(100);
			repo.RBC_Portal.SearchRequest.NasReqNum.PressKeys(varNasNbr);
			repo.RBC_Portal.SearchRequest.ViewGroupReq.Click();
			repo.RBC_Portal.SearchRequest.SearchSubmitBtn.Click();
			Delay.Milliseconds(100);
			
			//Post General Comment
			Comment postComment = new Comment();
			postComment.GeneralComment();
						
			
			//Close Browser
			//Host.Local.KillBrowser("IE");
			Delay.Milliseconds(200);
			
		}
	}
}
