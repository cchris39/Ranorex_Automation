/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-12
 * Time: 4:28 PM
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

namespace Scotia_Portal
{
	/// <summary>
	/// Description of Comment.
	/// </summary>
	[TestModule("9B2B9D1E-0310-410C-A38A-62AE53A22807", ModuleType.UserCode, 1)]
	public class Comment : ITestModule
	{
		
		
		public static Scotia_PortalRepository repo = Scotia_PortalRepository.Instance;
		
			static Comment instance = new Comment();
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Comment()
		{
			// Do not delete - a parameterless constructor is required!
		}
		#region	Varilables
		
		string _varNasNbr = "";
		[TestVariable("4152D611-8849-46DE-AB20-C8EC6C18295B")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		
		public void gereralComment(string comment)
		{
			repo.DomScotia.Comment.CommentOnRequestBtn.Click();
			repo.DomScotia.Comment.GeneralComment.Click();
			Delay.Milliseconds(100);
			repo.DomScotia.Comment.CommentText.PressKeys(comment);
			repo.DomScotia.Comment.SubmitCommentBtn.Click();
			Delay.Milliseconds(100);
		}
		
		public void logInternalComment(string comment)
		{
			repo.DomScotia.Comment.CommentOnRequestBtn.Click();
			repo.DomScotia.Comment.LogInternal.Click();
			Delay.Milliseconds(100);
			repo.DomScotia.Comment.CommentText.PressKeys(comment);
			repo.DomScotia.Comment.SubmitCommentBtn.Click();
			Delay.Milliseconds(100);
		}
		
		public void reportCommentStatus(string comment)
		{
			string commID = repo.DomScotia.Comment.CommentID.InnerText.Trim();
			string postComment = repo.DomScotia.Comment.CommentSubmitted.InnerText.Trim();
			string commentTime = repo.DomScotia.Comment.CommentSubmittedTime.InnerText.Trim();
			
			Report.Log(ReportLevel.Success, "Validation", "Comment successfully submitted.");
			Validate.Exists(repo.DomScotia.Comment.DivTagCommentsSuccessfullySubmitted);
			Validate.AreEqual(postComment, comment);
			
		}
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			const string genComment = "This is a Scotia General comment";
			const string logComment = "This is a Scotia Log Internal comment";
			
			/*
			Login UsrLogin = new Login();
			UsrLogin.lauchScotia();
			UsrLogin.UserLogin(uid, pwd);
			Delay.Milliseconds(100);
			*/
			
			//Post General Comment
			SearchRequest searchNas = new SearchRequest();
			searchNas.search();
			searchNas.findByNasNbr(varNasNbr);
			Delay.Milliseconds(100);
			
			gereralComment(genComment);
			reportCommentStatus(genComment);
			Delay.Milliseconds(100);
			
			//Log Internal Comment
			searchNas.search();
			searchNas.findByNasNbr(varNasNbr);
			Delay.Milliseconds(100);
			logInternalComment(logComment);
			reportCommentStatus(logComment);
			Delay.Milliseconds(100);
			
			Delay.Milliseconds(200);
			
			
		}
	}
}
