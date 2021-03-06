/*
 * Created by Ranorex
 * User: an.zhou
 * Date: 04/04/2016
 * Time: 11:43 AM
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
	/// Description of UserCodeModule1.
	/// </summary>
	[TestModule("22BA6AED-0B31-4410-A397-F8DD9EE488F1", ModuleType.UserCode, 1)]
	public class RequestService : ITestModule
	{
		
		public static RBC_TestRepository repo = RBC_TestRepository.Instance;
		
		static RequestService instance = new RequestService();
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		
		#region Varilables
		
		#endregion
			
		public RequestService()
		{
			// Do not delete - a parameterless constructor is required!
		}

		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		
		public void selectRequest(string varCategory, string varServiceType)
		{
			if (varCategory == "Linx")
			{
				repo.RBC_Portal.reqService.selectService.linxApplication.Click();
				if (varServiceType == "Standalone MarketRent")
				{
					repo.RBC_Portal.reqService.selectService.linxMarketRent.Click();
					Delay.Milliseconds(200);
				}else if (varServiceType == "Progress Inspection")
				{
					repo.RBC_Portal.reqService.selectService.linxPI.Click();
					Delay.Milliseconds(200);
				}
			// End of Linx Category
			}else if (varCategory == "Homebase")
			{
				repo.RBC_Portal.reqService.selectService.homebaseApplication.Click();
				if(varServiceType == "Standard")
				{
					repo.RBC_Portal.reqService.selectService.homebaseStdValuation.Click();
					Delay.Milliseconds(200);
				}else if (varServiceType == "New Construction")
				{
					repo.RBC_Portal.reqService.selectService.homebaseNewConstruction.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Major Renovation")
				{
					repo.RBC_Portal.reqService.selectService.homebaseMajorRenovation.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Market Rent With Appraisal")
				{
					repo.RBC_Portal.reqService.selectService.homebaseMarketRentWithAppraisal.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Standalone MarketRent")
				{
					repo.RBC_Portal.reqService.selectService.homebaseMarketRent.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "RuralEstate Residential")
				{
					repo.RBC_Portal.reqService.selectService.homebaseRuralEstate.Click();
					repo.RBC_Portal.reqService.selectService.homebseRuralResidential.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "RuralEstate Agricultural")
				{
					repo.RBC_Portal.reqService.selectService.homebaseRuralEstate.Click();
					repo.RBC_Portal.reqService.selectService.homebaseRuralAgricultural.Click();
					Delay.Milliseconds(100);
				}
			//------------------End of Homebase Category---------------------------------------------------------
			} else if (varCategory == "CASPER")                     
			{
				repo.RBC_Portal.reqService.selectService.casperApplication.Click();
				if(varServiceType == "Standard")
				{
					repo.RBC_Portal.reqService.selectService.CasperStdValuation.Click();
					Delay.Milliseconds(200);
				}else if (varServiceType == "New Construction")
				{
					repo.RBC_Portal.reqService.selectService.casperNewConstruction.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Major Renovation")
				{
					repo.RBC_Portal.reqService.selectService.casperMajorRenovation.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Market Rent With Appraisal")
				{
					repo.RBC_Portal.reqService.selectService.casperMarketRenyWithAppraisal.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "Standalone MarketRent")
				{
					repo.RBC_Portal.reqService.selectService.casperMarketRent.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "RuralEstate Residential")
				{
					repo.RBC_Portal.reqService.selectService.casperRuralEstate.Click();
					repo.RBC_Portal.reqService.selectService.casperRuralResidential.Click();
					Delay.Milliseconds(100);
				}else if (varServiceType == "RuralEstate Agricultural")
				{
					repo.RBC_Portal.reqService.selectService.casperRuralEstate.Click();
					repo.RBC_Portal.reqService.selectService.casperRuralAgricultural.Click();
					Delay.Milliseconds(100);
				}
				//------------------------- End of CASPER Category -----------------------------------------
			} else if (varCategory == "Other" && varServiceType == "Progress Inspection")
				{
					repo.RBC_Portal.reqService.selectService.PI_NewConstructionApplication.Click();
					Delay.Milliseconds(100);
				}
			else if (varCategory == "Other" && varServiceType == "Appraisal Update")
				{
					repo.RBC_Portal.reqService.selectService.appraisalUpdateApplicatin.Click();
					Delay.Milliseconds(100);
				}
			//End of Other Category
		} //End of SelectRequest
		
		
		public void linxRequest(string varServiceType, string postalCode, string streetNum, string streetName, string streetType, 
		                        string loanType, string appPrice, string appMtAmount, string appFstName, string appLstName, 
		                        string appRefNbr, string transit, string language, string spText, string oriNbr)
		{
				if (varServiceType == "Standalone MarketRent")
				{
					repo.RBC_Portal.reqService.selectService.linxMarketRent.Click();
					
					repo.RBC_Portal.reqService.requestForm.postalCode.PressKeys(postalCode);
					repo.RBC_Portal.reqService.requestForm.postalCode.PressKeys("{Tab}");
					Delay.Milliseconds(100);
			
					repo.RBC_Portal.reqService.requestForm.streetNbr.PressKeys(streetNum);
					repo.RBC_Portal.reqService.requestForm.streetName.PressKeys(streetName);
					repo.RBC_Portal.reqService.requestForm.streetType.PressKeys(streetType);
			
					Ranorex.SelectTag reqLoanType = "/dom[@domain='uattest.nas.com']//select[@id='field_value_47_copy']";
					reqLoanType.TagValue = loanType;
					reqLoanType.Click();
					Delay.Milliseconds(100);
			
					repo.RBC_Portal.reqService.requestForm.LinxPurchasePrice.PressKeys(appPrice);
					repo.RBC_Portal.reqService.requestForm.LinxMortgAmount.PressKeys(appMtAmount);
					Delay.Milliseconds(100);
						
					repo.RBC_Portal.reqService.requestForm.LinxClientFstName.PressKeys(appFstName);
					repo.RBC_Portal.reqService.requestForm.LinxClientLstName.PressKeys(appLstName);
					repo.RBC_Portal.reqService.requestForm.LinxApplicationNbr.PressKeys(appRefNbr);
			
					repo.RBC_Portal.reqService.requestForm.LinxTransitNbr.PressKeys(transit);
					Delay.Milliseconds(200);
			
					Ranorex.SelectTag reqLanguage = "/dom[@domain='uattest.nas.com']//select[@id='field_value_21_copy']";
					//reqLanguage.TagValue = language;
					//reqLanguage.Click();
					repo.RBC_Portal.reqService.requestForm.LinxSpText.PressKeys(spText);
					Delay.Milliseconds(100);
			
				}else if (varServiceType == "Progress Inspection")
				{
					repo.RBC_Portal.reqService.selectService.linxPI.Click();
					repo.RBC_Portal.reqService.requestForm.LinxOrgReqNbr_PI.PressKeys(oriNbr);
					repo.RBC_Portal.reqService.requestForm.Linx_OriNbrSearchBtn.Click();
					Delay.Milliseconds(200);
				}
					repo.RBC_Portal.reqService.requestForm.previewRequestBtn.Click();
					Delay.Milliseconds(200);
			
					if (repo.RBC_Portal.reqService.requestForm.ConfirmPreviewPopupInfo.Exists())
					{
					repo.RBC_Portal.reqService.requestForm.SubmitRequest.Click();
					Delay.Milliseconds(300);
					}else
					{Report.Log(ReportLevel.Warn,"Warning", "No appraisal request preview information present, please check.");}
				
		} //End of linxRequest
		
		public void portalRequest(string varServiceType, string postalCode, string streetNum, string streetName, string streetType,
		                          string loanType, string appPrice, string appMtAmount, string appFstName, string appLstName, string appRefNbr, string purchaseType,
		                          string transit, string clientScore, string contactFstName, string contactLstName,string tel1, string tel2, string tel3, 
		                          string language, string spText, string oriNbr, string builder, string builderCode, string siteName,
		                          string propertyType, string tenure, string mortgageType, string programType, string appraisalType)
		{
						
						//Property Address
						repo.RBC_Portal.reqService.requestForm.postalCode.PressKeys(postalCode);
						repo.RBC_Portal.reqService.requestForm.postalCode.PressKeys("{Tab}");
						Delay.Milliseconds(100);
					
						repo.RBC_Portal.reqService.requestForm.streetNbr.PressKeys(streetNum);
						repo.RBC_Portal.reqService.requestForm.streetName.PressKeys(streetName);
						repo.RBC_Portal.reqService.requestForm.streetType.PressKeys(streetType);
					
						//Transcation Details
						if (varServiceType == "Standard")
						{
						repo.RBC_Portal.reqService.requestForm.propertyType.TagValue = propertyType;
						repo.RBC_Portal.reqService.requestForm.propertyType.Click();
						repo.RBC_Portal.reqService.requestForm.tenure.TagValue = tenure;
						repo.RBC_Portal.reqService.requestForm.tenure.Click();
						repo.RBC_Portal.reqService.requestForm.LoanType.TagValue = loanType;
						repo.RBC_Portal.reqService.requestForm.LoanType.Click();
						repo.RBC_Portal.reqService.requestForm.Price.PressKeys(appPrice);
						repo.RBC_Portal.reqService.requestForm.MortgAmount.PressKeys(appMtAmount);
						
						repo.RBC_Portal.reqService.requestForm.mortgageType.TagValue = mortgageType;
						repo.RBC_Portal.reqService.requestForm.mortgageType.Click();
						repo.RBC_Portal.reqService.requestForm.programType.TagValue = programType;
						repo.RBC_Portal.reqService.requestForm.programType.Click();
						repo.RBC_Portal.reqService.requestForm.environmentalHazd.Click();
						Delay.Milliseconds(100);
						} else if (varServiceType == "New Construction" || varServiceType == "Major Renovation")
						{
						repo.RBC_Portal.reqService.requestForm.PurchaseType.TagValue = purchaseType;
						repo.RBC_Portal.reqService.requestForm.ConstructPrice.PressKeys(appPrice);
						repo.RBC_Portal.reqService.requestForm.ConstructMortgageAmount.PressKeys(appMtAmount);
						repo.RBC_Portal.reqService.requestForm.BuilderType.TagValue = builder;
						repo.RBC_Portal.reqService.requestForm.BuilderCode.TagValue = builderCode;
						repo.RBC_Portal.reqService.requestForm.SiteProjectName.PressKeys(siteName);
						Delay.Milliseconds(100);
						} else if (varServiceType == "Standalone MarketRent" || varServiceType == "Market Rent With Appraisal" )
						{
						repo.RBC_Portal.reqService.requestForm.MarketRentPropertyType.TagValue = propertyType;
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralLoanType.TagValue = loanType;
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralLoanType.Click();
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralPrice.PressKeys(appPrice);
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralMortgAmount.PressKeys(appMtAmount);
						} else if (varServiceType == "RuralEstate Agricultural")	
						{
							repo.RBC_Portal.reqService.requestForm.RuralAppraisalType.TagValue = appraisalType;
							repo.RBC_Portal.reqService.requestForm.MarketRentRuralLoanType.TagValue = loanType;
							repo.RBC_Portal.reqService.requestForm.MarketRentRuralPrice.PressKeys(appPrice);
							repo.RBC_Portal.reqService.requestForm.MarketRentRuralMortgAmount.PressKeys(appMtAmount);
						} else
						{
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralLoanType.TagValue = loanType;
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralLoanType.Click();
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralPrice.PressKeys(appPrice);
						repo.RBC_Portal.reqService.requestForm.MarketRentRuralMortgAmount.PressKeys(appMtAmount);
						}
						Delay.Milliseconds(100);
						
						//LOAN APPLICATION
						if (varServiceType == "Standalone MarketRent")       //Omit Final Cient Score and Input Cient name in different object
						{
						repo.RBC_Portal.reqService.requestForm.LinxClientFstName.PressKeys(appFstName);
						repo.RBC_Portal.reqService.requestForm.LinxClientLstName.PressKeys(appLstName);
						repo.RBC_Portal.reqService.requestForm.LinxApplicationNbr.PressKeys(appRefNbr);
						repo.RBC_Portal.reqService.requestForm.LinxTransitNbr.PressKeys(transit);
						repo.RBC_Portal.reqService.requestForm.LinxSpText.PressKeys(spText);
						}
						else
						{
						repo.RBC_Portal.reqService.requestForm.ClientFstName.PressKeys(appFstName);
						repo.RBC_Portal.reqService.requestForm.ClientLstName.PressKeys(appLstName);
						repo.RBC_Portal.reqService.requestForm.casperNbr.PressKeys(appRefNbr);
						repo.RBC_Portal.reqService.requestForm.transitNbr.PressKeys(transit);
						repo.RBC_Portal.reqService.requestForm.finalClientScore.PressKeys(clientScore);
						
						// Contact Name
						repo.RBC_Portal.reqService.requestForm.ContactFstName.PressKeys(contactFstName);
						repo.RBC_Portal.reqService.requestForm.ContactLstName.PressKeys(contactLstName);
						repo.RBC_Portal.reqService.requestForm.contact1stNbr_1.PressKeys(tel1);
						repo.RBC_Portal.reqService.requestForm.contact1stNbr_2.PressKeys(tel2);
						repo.RBC_Portal.reqService.requestForm.contact1stNbr_3.PressKeys(tel3);
						Delay.Milliseconds(200);
						}
						
						//Information to Appraiser
						repo.RBC_Portal.reqService.requestForm.languageOption.TagValue = language;
						
						if (varServiceType == "Standard" || varServiceType == "RuralEstate Agricultural" || varServiceType == "RuralEstate Residential")
						{
							repo.RBC_Portal.reqService.requestForm.stdSpdirectionText.PressKeys(spText);
						}else if (varServiceType != "Standalone MarketRent")
						{
							repo.RBC_Portal.reqService.requestForm.spDirectionText.PressKeys(spText);
						}
						
						Delay.Milliseconds(100);
											
				//---------------------------------------------------------------
						repo.RBC_Portal.reqService.requestForm.previewRequestBtn.Click();
						Delay.Milliseconds(600);
			
					if (repo.RBC_Portal.reqService.requestForm.ConfirmPreviewPopupInfo.Exists())
					{
						repo.RBC_Portal.reqService.requestForm.SubmitRequest.Click();
						Delay.Milliseconds(300);
					}else
					{	Report.Log(ReportLevel.Warn,"Warning", "No appraisal request preview information present, please check.");}
				
		} //End of portalRequest
		
		public void portalOtherRequest(string varServiceType, string oriNbr)
		{
			if (varServiceType == "Progress Inspection")
			{
				repo.RBC_Portal.reqService.requestForm.OrgReqNbr_PI.PressKeys(oriNbr);
				repo.RBC_Portal.reqService.requestForm.PI_OriNbrSearchBtn_other.Click();
				Delay.Milliseconds(100);
			}else if (varServiceType == "Appraisal Update")
			{
				repo.RBC_Portal.reqService.requestForm.OrgReqNbr_AppraisalUpdate.PressKeys(oriNbr);
				repo.RBC_Portal.reqService.requestForm.searchBtn_AppraisalUpdate.Click();
				Delay.Milliseconds(100);
			}
			
			//---------------------------------------------------------------
						repo.RBC_Portal.reqService.requestForm.previewRequestBtn.Click();
						Delay.Milliseconds(600);
			
					if (repo.RBC_Portal.reqService.requestForm.ConfirmPreviewPopupInfo.Exists())
					{
						repo.RBC_Portal.reqService.requestForm.SubmitRequest.Click();
						Delay.Milliseconds(300);
					}else
					{	Report.Log(ReportLevel.Warn,"Warning", "No appraisal request preview information present, please check.");}
		}
		//End of portalOtherRequest
				
		//-------------Since the order submit confirmation page is not identical based on the service type option selected, thus the below part is needed to differentiate the senario ---------//
		public string GetOtherRequestNbr()        // Get other request Nas number after order submit
		{
			string nasNbr = repo.RBC_Portal.ATagNasNbr.InnerText.Trim();
			return nasNbr;
		}

		public string GetStdRequestNbr()        // Get standard request Nas number after order submit
		{
			string nasNbr = repo.RBC_Portal.BTagNasNbr.InnerText.Trim();
			return nasNbr;
		}
		
		public string GetOtherServiceType( )    // Get other request service type after order submit
		{
			string orderType = repo.RBC_Portal.ThePropertyValuationOrderedIs_other.InnerText.Trim();
			string[] s = orderType.Split(':');
			string serviceType = s[1].Trim();
						
			return serviceType;
		}
		public string GetStdServiceType()    // Get standard request service type after order submit
		{
			string orderType = repo.RBC_Portal.ValuationHasBeenLogg_std.InnerText.Trim();
			string[] s = orderType.Split(':');
			string serviceType = s[2].Trim();
						
			return serviceType;
		}
		
		//------------------------ End of info for order confirmation page --------------------------------------//
	
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
		}
	}
}
