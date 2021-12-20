/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 2016-05-03
 * Time: 4:17 PM
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
	/// Description of QualityComment.
	/// </summary>
	[TestModule("3FA4A03C-C51A-4216-BC71-D6FEF8875A81", ModuleType.UserCode, 1)]
	public class QualityComment : ITestModule
	{
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static QualityComment instance = new QualityComment();
				
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public QualityComment()
		{
			// Do not delete - a parameterless constructor is required!
		}
		
		#region Varilables
		
		string _varUid = "";
		[TestVariable("19162D13-3153-4D5B-AAF1-E9BD4DD6A592")]
		public string varUid
		{
			get { return _varUid; }
			set { _varUid = value; }
		}
		
		string _varNasNbr = "";
		[TestVariable("E37AFDE4-2DCF-4A94-BDA7-FB02F7413BF4")]
		public string varNasNbr
		{
			get { return _varNasNbr; }
			set { _varNasNbr = value; }
		}
		
		#endregion
		
		private string uid = "rbc@nationwideappraisals.com";
    	private string pwd = "XXXXXX";
    	private string url = "https://uattest.nas.com/NAS/rbcpvpnas";
    	
    	private const string val_CR1 = "CR01";
    	private const string val_CR2 = "CR02";
		private const string val_CR3 = "CR03";
    	
		private const string cr1Txt= "Missing Character 'A' in the street number for this property.";
		private const string cr2Txt= "Appraiser don't have required certification.";
		private const string cr3Txt = "Final report does not meet lender requirement.";
		
		public void postQualityComment(string commValue, string commentTxt, string nbr)
    	{
    		repo.RBC_Portal.QualityValueAppeal.QualityIssueSelected.Click();
    		Delay.Milliseconds(100);
    		repo.RBC_Portal.QualityValueAppeal.QualityIssueWithDropDownCategoryList.TagValue = commValue;
    		Delay.Milliseconds(100);
    		
    		if(commValue == "CR03")
    		{	
    			repo.RBC_Portal.QualityValueAppeal.MissLocationMapCheckBox.Click();
    			repo.RBC_Portal.QualityValueAppeal.MissCompPhotoCheckBox.Click();
    			repo.RBC_Portal.QualityValueAppeal.MissInteriorPhotoCheckBox.Click();
    			
    		}
    		repo.RBC_Portal.QualityValueAppeal.CommentArea.PressKeys(commentTxt);
    		repo.RBC_Portal.QualityValueAppeal.SubmitComment.Click();
    		Delay.Milliseconds(300);
    		
    		string postTxt = repo.RBC_Portal.QualityValueAppeal.CommentsSubmitted.InnerText.Trim();
    		string commentID = repo.RBC_Portal.QualityValueAppeal.CommentID.InnerText.Trim();
    		string[] chgStatus = repo.RBC_Portal.QualityValueAppeal.CurrentStatusUnderReview.InnerText.Trim().Split (':');
    		string curStaus = chgStatus[1];
    		
    		
    		Report.Success("Quality Comment: " + "\"" + commentTxt + "\"" + " successfully submitted for " + nbr + ". Comment ID: " + commentID );
    		Report.Info(nbr + " report current status: " + curStaus);
    		Validate.Exists(repo.RBC_Portal.QualityValueAppeal.QualityCommentsSuccessfulSubmit);
    		//Validate.AreEqual(commentTxt,postTxt);
    	
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
						
			string sql = "SELECT app_request_nbr FROM app_request a where username = '" + varUid + "' and status = 'completed' and service_residential = 'Full-Service' order by app_request_nbr asc limit 1;";
			DataTable dt = QueryDB.RunQuery(sql);
			
			string nbr = dt.Rows[0][0].ToString().Trim();
			
			varNasNbr = nbr;
			
			SearchRequest searchNbr = new SearchRequest();
			searchNbr.findByNasNbr(nbr);
			
			//Select request to submit quality comment
			repo.RBC_Portal.SearchRequest.SelectRadioRow.Click();
    		Delay.Milliseconds(100);
    		repo.RBC_Portal.Main.CommentValueAppeal.Click();
    		Delay.Milliseconds(200);
    		
			postQualityComment(val_CR1, cr1Txt, nbr);
			Delay.Milliseconds(100);
			
			postQualityComment(val_CR2, cr2Txt, nbr);
			Delay.Milliseconds(100);
			
			postQualityComment(val_CR3, cr3Txt, nbr);
			Delay.Milliseconds(100);
			
		}
	}
}
