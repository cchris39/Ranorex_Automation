/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-03
 * Time: 4:20 PM
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
using MySql.Data.MySqlClient;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace RBC_Test
{
	/// <summary>
	/// Description of ValueAppeal.
	/// </summary>
	[TestModule("B80AC288-9CAF-4E27-A20A-2CE2847628C9", ModuleType.UserCode, 1)]
	public class ValueAppeal : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		public static Dom_AppraiserSanityTestRepository DomRepo = Dom_AppraiserSanityTestRepository.Instance;
		
		static ValueAppeal instance = new ValueAppeal();
		
		#region Varilables
		
		string _varUid = "";
		[TestVariable("26F6D2AA-E4D3-4064-B775-532F005BEB0E")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		
		string _varNasNbr = "";
		[TestVariable("DD84F32B-E74C-4206-A2D0-ADB5EF87B5FE")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		private string FilePath = @"c:\Upload_Files\SupportDocument.pdf\";
		private string uid = "testrbc@nas.com";
    	private string pwd = "xxxxxx";
    	private string url = "https://uattest.nas.com/NAS/rbcpvpnas";
    	
		private const string disagree = "Upon reviewed, value appeal validate, additional appraisals is warranted.";
		private const string agree = "Upon reviewed, I agree to the appraised value.";
		private const string app_Txt1 = "Different opinion on the value for this report, escalate to an Independent Review Appraiser.";
		
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public ValueAppeal()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		public void appealValue(string commTxt, string path, string agreeValue)
		{
			repo.RBC_Portal.QualityValueAppeal.ValueAppealEscalation.Click();
			Delay.Milliseconds(100);
						
			repo.RBC_Portal.QualityValueAppeal.SubmitComment.Click();
			
			//Verify reasoning pop op message
				if (repo.MessageFromWebpage.Text65535Info.Exists())
					{
						repo.MessageFromWebpage.ButtonOK.Click();
						Report.Log(ReportLevel.Info, "Value appeal reasoning comment pop op message verified.");
					}else
					{	Report.Log(ReportLevel.Warn, "Warning", "Value appeal reasoning comment pop op message not present, please check.");}
			
			
			if (agreeValue == "Agree")
			{	//Upload sales manger approval copy
				repo.RBC_Portal.QualityValueAppeal.SalesManagerApprovalRequired.DoubleClick();
				Delay.Milliseconds(100);
			}else if (agreeValue == "Disagree")
			{
				repo.RBC_Portal.QualityValueAppeal.RVPApprovalRequired.DoubleClick();
				Delay.Milliseconds(100);
			}
			
			repo.RBC_Portal.Main.FilePath.DoubleClick();
			Delay.Milliseconds(200);
			
			//repo.AdditionalDocumentsUploadInternetE.Self.Maximize();
			
			repo.ChooseFileToUpload.FilePath.PressKeys(path);
			repo.ChooseFileToUpload.ButtonOpen.Click();
			Delay.Milliseconds(200);
			repo.RBC_Portal.Main.FileUpload.Click();
			Delay.Milliseconds(600);
						
			Validate.Exists(repo.RBC_Portal.QualityValueAppeal.BigTagSalesManagerApprovedRequiredD);
			Report.Log(ReportLevel.Success, "Validation", "Sales Manager Approval document successfully uploaded");
						
			repo.AppraisalCommentsInternetExplorer.Self.Close();
			Delay.Milliseconds(400);
			
			//Input value appeal comment text
			repo.RBC_Portal.QualityValueAppeal.CommentArea.PressKeys(commTxt);
			repo.RBC_Portal.QualityValueAppeal.SubmitComment.Click();
			Delay.Milliseconds(300);
			
			Validate.Exists(repo.RBC_Portal.Main.CommentsSuccessfulSubmit);
			Report.Log(ReportLevel.Success, "Validation", "Request has been escalated to an indepent review appraiser. Comment submitted.");
			
		}
		
		public void reviewerComment(string agreeValue, string path, string nbr)
		{
			DomRepo.DomNasHome.MenuDisplay.AdminCommentOnRequest.Click();
			Delay.Milliseconds(200);
			
			if (agreeValue == "Agree")
			{
				DomRepo.DomNasHome.MenuDisplay.AgreesToAppraisedValue.Click();
				DomRepo.DomNasHome.MenuDisplay.ATagUploadReviewSummaryReport.Click();
				Delay.Milliseconds(300);
				
				DomRepo.DomNasHome.SummaryReportUploadPath.DoubleClick();
				Delay.Milliseconds(300);
				
				DomRepo.ChooseFileToUpload.FileName.PressKeys(path);
				DomRepo.ChooseFileToUpload.ButtonOpen.Click();
				Delay.Milliseconds(300);
				DomRepo.DomNasHome.Upload.Click();
				Delay.Milliseconds(600);
				
				Validate.Exists(DomRepo.DomNasHome.BigTagSummaryReportHaveBeenUploaded);
				Report.Log(ReportLevel.Success, "Success", "Summary Report has been upload successfully for " + nbr);
				
				DomRepo.AppraisalCommentsInternetExplorer.Self.Close();
				Delay.Milliseconds(200);
				
				DomRepo.DomNasHome.MenuDisplay.CommentText.PressKeys(agree);
				Delay.Milliseconds(300);
				DomRepo.DomNasHome.MenuDisplay.SubmitComment.Click();
				
			}else if (agreeValue == "Disagree")
			{
				DomRepo.DomNasHome.MenuDisplay.DoesNotAgreeToAppraisedValue.Click();
				DomRepo.DomNasHome.MenuDisplay.CommentText.PressKeys(disagree);
				DomRepo.DomNasHome.MenuDisplay.SubmitComment.Click();
				Delay.Milliseconds(300);
			}
				
				Validate.Exists(DomRepo.DomNasHome.MenuDisplay.CommentsWereSuccessfullySubmitted);
				Report.Log(ReportLevel.Info, "Information", "Review Appraiser comment successfully submitted for " + nbr );
				Delay.Milliseconds(100);
				
			//Verify status changed to 'Completed'	
			DomRepo.DomNasHome.SearchFilter.Click();
			Delay.Milliseconds(200);
			
			DomRepo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(nbr);
			DomRepo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			string reqStatus = DomRepo.DomNasHome.MenuDisplay.curStatus.InnerText.Trim();
			Validate.AreEqual(reqStatus, "Completed");
			Report.Log(ReportLevel.Info, "Validation", nbr + " request current status: " + reqStatus);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
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
			
			//RBC user log in
			Login userLogin = new Login();
			userLogin.login(uid, pwd, url);
						
			//Agree work flow
			string reviewValue = "Agree"; 
			
			//string sql1 = "SELECT app_request_nbr FROM app_request a where username = '" + varUid + "' and status = 'completed' and service_residential = 'Full-Service' order by app_request_nbr desc limit 1;";
			//DataTable dt1 = QueryDB.RunQuery(sql1);
			
			//string nbr = dt1.Rows[0][0].ToString().Trim();
			
			string nbr = varNasNbr;
			
			SearchRequest searchNbr = new SearchRequest();
			searchNbr.findByNasNbr(nbr);
			
			//Select request to submit value appeal
			repo.RBC_Portal.SearchRequest.SelectRadioRow.Click();
    		Delay.Milliseconds(100);
    		repo.RBC_Portal.Main.CommentValueAppeal.Click();
    		Delay.Milliseconds(200);
			
    		appealValue(app_Txt1, FilePath, reviewValue);
    		
    		//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
    		
			
			//Review Appraiser start his flow
			Host.Local.OpenBrowser("http://uattest.nas.com/NAS", "IE", "", false, false, false, false, false);
			Delay.Milliseconds(400);
			
			DomRepo.DomNasHome.Userid.PressKeys("admin");
			DomRepo.DomNasHome.Password.PressKeys("Tester#1");
			DomRepo.DomNasHome.Submit.Click();
			Delay.Milliseconds(300);
			
			DomRepo.DomNasHome.SearchFilter.Click();
			Delay.Milliseconds(200);
			
			DomRepo.DomNasHome.MenuDisplay.NasReqNum.PressKeys(nbr);
			DomRepo.DomNasHome.MenuDisplay.SearchSubmit.Click();
			Delay.Milliseconds(200);
			
			//Verify request current status changed to 'Under Review'
			string reqStatus = DomRepo.DomNasHome.MenuDisplay.curStatus.InnerText.Trim();
			Validate.AreEqual(reqStatus, "Under Review");
			Report.Log(ReportLevel.Info, "Validation", nbr	+ " request current status: " + reqStatus);
			
			//Review Appraiser Post Comment
			reviewerComment(reviewValue, FilePath, nbr);
			
			
			//RBC HEF User flow
			userLogin.login("anhua.zhou@nationwideappraisals.com", "Tester#9", url);
			Delay.Milliseconds(100);
			
			string useOriginal = "The reports have been reviewed, please use the appraised value on the original report, ensure all original approval conditions have been met.";
			string hefTxt = "Please use the original report.";
	
			searchNbr.findByNasNbr(nbr);
			
			repo.RBC_Portal.Main.CommentValueAppeal.Click();
			repo.RBC_Portal.QualityValueAppeal.SendCommentToNas.Click();
			repo.RBC_Portal.QualityValueAppeal.UseOriginalReport.Click();
			
			Delay.Milliseconds(400);
			
			Validate.Exists(repo.RBC_Portal.Main.CommentsSuccessfulSubmit);
			Report.Log(ReportLevel.Info, "Validation", "RBC HEF user comment successfullu submitted: " + "\"" + useOriginal + "\"");
			Delay.Milliseconds(200);
			
			//Close Browser
			Host.Local.KillBrowser("IE");
			Delay.Milliseconds(300);
    		
		}
	}
}
