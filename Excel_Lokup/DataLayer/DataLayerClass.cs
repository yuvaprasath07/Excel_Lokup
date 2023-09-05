using Entity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DataLayer
{
    public class DataLayerClass : IDatalayer
    {
        public readonly LokupDbContext _dbContext;
        public DataLayerClass(LokupDbContext dbContext)
        {
            _dbContext = dbContext;
        }
        public object GetExcelLokup()
        {

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {

                var sheet1 = package.Workbook.Worksheets.Add("BatchEnrollment");
                string[] columnHeaders = {
            "contractor id", "Customer code", "Public office/private sector*",
            "Customer type*", "Corporate Name*", "Corporate Name Kana*",
            "contract_postal code*", "Contract_Address 1 (Prefecture + City)*",
            "Contract_Address 2 (Municipality)*", "Contract _ chome",
            "contract address1", "contract number1", "contract_building name",
            "contract building", "contract room1", "Contract_person in charge name*",
            "Contract_person in charge name (Kana)*", "Contract_Department/Affiliation*",
            "contract_phone number*", "Contract_FAX", "E Mail Address*",
            "SalesforceID (issue)", "Project No.", "Project Title",
            "SalesforceID (account)", "contract_company name*",
            "Contract delivery_Company name Kana*",
            "Contract delivery_person in charge name*",
            "Contract delivery_person in charge name (Kana)*",
            "Contract delivery_Department/Affiliation*",
            "Contract delivery_postal code*",
            "Contract delivery address 1 (prefecture + city)*",
            "Contract delivery_address 2 (town/village)*",
            "contract _ chome", "contract address2", "contract number2",
            "Contract delivery_building name", "contract delivery _ building",
            "contract room2", "contract_telephone number*", "Contract delivery_FAX",
            "contract_email address*", "billing address", "Contract_Company name*",
            "Contract_Company Name (kana)*", "Request_Department/Affiliation**",
            "request_person in charge name*", "Request_person in charge name (kana)*",
            "request_postal code*", "Request_Address 1 (Prefecture + City)*",
            "Request_Address 2 (Municipality)", "Request _ chome","Contract billing address ","contract number3",
            "Contract_Building name","building","billing room","request_phone number*",
            "Request_FAX","Request_E Mail Address*","Payment Method*","Billing method*",
            "Newly established","New continuation category","Enrollment R 1) 1*",
            "Registered Kana*","Service_zip code*","Service_address 1 (prefecture + city)*",
            "Service_ Address 2 (town/village)*","Service _ chome","service address",
            "Service_go","_ building name","service building","service room","SPID*",
            "area*","Name of facility*","Quotation No*","SalesforceID (demand base)",
            "Scheduled supply start date*","Weighing day*","Contract Term*","Main contract system*",
            "main fee structure*","load pattern*","Contracts1*","Utility*","Rate Menu","Salesforce ID (plan)",
            "contract binding period*","Contract automatic renewal category*","Estimate*","Tax Rate*","Contract start date*",
            "Contract end date*","Contract change date","Old invoice Contract name",
            "Current Supply Class","Current contract power","Current Retailer Name","Current Retailer Customer Number",
            "Current base contract power","Current base contract supply company name","Current base contract customer number","demand window_company name*",
            "Kisumo_Company Name Kana*","Demand window_person in charge name*","Demand window_person in charge name in kana*","Sales window_department/affiliation*","demand window_postal code*",
            "Demand window_address 1 (prefecture + city)*","Demand window_address 2 (town/village)*","demand window _ chome","demand window address","demand window _ number","Demand window_Building name",
            "demand window _ building","demand window _ room","demand window_telephone number*","Demand window_FAX","demand window_email address*","technique_company name ","Technique_Company Name Kana","Technique_person in charge name",
            "Tech_Person in charge Name Kana","Technology_Department/Affiliation",
            "technique_phone number","Renewable energy surcharge exemption target category*","Renewable energy surcharge exemption rate","Renewable energy surcharge exemption start date",
            "Renewable energy surcharge exemption end date","Environment menu target category*","initiative","environmental value","Power configuration","Certificate usage","Renewable energy ratio","Contracts2","Renewable energy supply start date","Renewable energy supply end date","main_supply voltage*",
            "Main_metering voltage*","Main_supply category*","Selected when main_partial supply","Main contract power*","main_base contract power","main_basic unit price*","Wheeling contract power","main_general unit price","main_weekday unit price","Main_daytime unit price","Main_weekend unit price","Main_night unit price","Main_holiday unit price",
            "Main_Heavy load unit price","main_peak unit price","Main_summer unit price","Main_other season unit price","Main_summer weekday unit price","Main_other season weekday unit price","Main_summer daytime unit price",
            "Main_other season daytime unit price","Main_summer holiday unit price","Main_other season holiday unit price","Spare line target classification*","Reserve line_contract power","Backup line_basic unit price","Spare line_supply voltage","Spare line_metering voltage","Standby power supply target category*",
            "Standby power source_Contract power","Standby power supply_basic charge unit price","Standby power source_supply voltage","Standby power supply_metering voltage","Self-supporting target classification*","Self-supplementary_reference power","Self-supplementary power contract","Self-supplementary power calculation method","Self-support_contract system",
            "Self-supplement_Monthly basic charge unit price","Private Supplement_Monthly basic charge unit price when not in use",
            "Self-supplementary _ regular summer unit price","Private Supplement_Irregular Summer Unit Price","Private Supplement_Regular Other Seasonal Unit Price","Self-supplementary_irregular other seasonal unit price","Fee unit price","Partition basic charge unit price","Customer Number","Billing Account Number",
        };

                for (int i = 0; i < columnHeaders.Length; i++)
                {
                    sheet1.Cells[1, i + 1].Value = columnHeaders[i];
                }

                var sheet2 = package.Workbook.Worksheets.Add("Lookups");

                sheet2.Cells["A1"].Value = "Self-supplementary power calculation method";
                sheet2.Cells["B1"].Value = "Payment Method";
                sheet2.Cells["C1"].Value = "Load pattern";
                sheet2.Cells["D1"].Value = "Contracts";
                sheet2.Cells["E1"].Value = "Current Supply Class";
                sheet2.Cells["F1"].Value = "Initiative";
                sheet2.Cells["G1"].Value = "Power configuration";
                sheet2.Cells["H1"].Value = "Certificate usage";
                sheet2.Cells["I1"].Value = "PowerSupplyContractTypes(contracts2)";
                sheet2.Cells["J1"].Value = "Main_supply category";
                sheet2.Cells["K1"].Value = "Customer Types";
                sheet2.Cells["L1"].Value = "Billing Method";
                sheet2.Cells["M1"].Value = "Utility";
                sheet2.Cells["N1"].Value = "Area";
                sheet2.Cells["O1"].Value = "Public office/private sector";
                sheet2.Cells["P1"].Value = "New continuation category";
                sheet2.Cells["Q1"].Value = "Main fee structure";
                sheet2.Cells["R1"].Value = "Estimate";
                sheet2.Cells["S1"].Value = "Newly established";
                sheet2.Cells["T1"].Value = "Contract automatic renewal category";
                sheet2.Cells["U1"].Value = "Renewable energy surcharge exemption target category";
                sheet2.Cells["V1"].Value = "Environment menu target category";
                sheet2.Cells["W1"].Value = "Selected when main_partial supply";
                sheet2.Cells["X1"].Value = "Spare line target classification";
                sheet2.Cells["Y1"].Value = "Standby power supply target category";
                sheet2.Cells["Z1"].Value = "Self-supporting target classification";

                var range = sheet2.Cells["A1:Z1"];
                range.Style.Font.Bold = true;
                var headerStyle = sheet1.Cells["A1:GH1"].Style;
                headerStyle.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                headerStyle.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                var Area = _dbContext.L_Area.ToList();
                var SelfSupplementary = _dbContext.L_SelfSupplementaryPowerCalculationMethod.ToList();
                var PaymentMethod = _dbContext.L_PaymentMethod.ToList();
                var Loadpattern = _dbContext.L_Loadpattern.ToList();
                var PowerSupplyContractTypes = _dbContext.L_PowerSupplyContractTypesContracts2.ToList();
                var Utility = _dbContext.L_Utility.ToList();
                var MainFeeStructure = _dbContext.L_MainFeeStructure.ToList();

                for (int i = 0; i < SelfSupplementary.Count; i++)
                {
                    sheet2.Cells[i + 2, 1].Value = SelfSupplementary[i].Description;

                }
                for (int i = 0; i < PaymentMethod.Count; i++)
                {
                    sheet2.Cells[i + 2, 2].Value = PaymentMethod[i].Description;

                }
                for (int i = 0; i < Loadpattern.Count; i++)
                {
                    sheet2.Cells[i + 2, 3].Value = Loadpattern[i].Description;

                }
                for (int i = 0; i < PowerSupplyContractTypes.Count; i++)
                {
                    sheet2.Cells[i + 2, 9].Value = PowerSupplyContractTypes[i].Description;

                }
                for (int i = 0; i < Utility.Count; i++)
                {
                    sheet2.Cells[i + 2, 13].Value = Utility[i].Description;

                }
                for (int i = 0; i < Area.Count; i++)
                {
                    sheet2.Cells[i + 2, 14].Value = Area[i].Description;

                }
                for (int i = 0; i < MainFeeStructure.Count; i++)
                {
                    sheet2.Cells[i + 2, 17].Value = MainFeeStructure[i].Description;

                }

                //Hotcore
                string[] contractsData = new string[] { "Actual demand", "Consultation", };
                string[] CurrentSupplyClass = new string[] { "Full supply", "Partial supply" };
                string[] Initiative = new string[] { "CDP", "SBT", "RE100" };
                string[] PowerConfiguration = new string[] { "With specific request", "Without specific request" };
                string[] CertificateUsage = new string[] { "With specific request", "Without specific request " };
                string[] Mainsupplycategory = new string[] { "Full supply", "Partial supply " };
                string[] CustomerTypes = new string[] { "Residential", "Commercial" };
                string[] BillingMethod = new string[] { "Web ", "Mail" };
                string[] PublicOfficeprivateSector = new string[] { "Public", "Private" };
                string[] NewContinuationCategory = new string[] { "New", "Continuous" };
                string[] MainFeeStructures = new string[] { "Seasonal ", " TOU", "Holiday heavy load", "Terms and Conditions", "Others" };
                string[] Estimate = new string[] { "New ", "Add base ", "Change unit price" };
                string[] NewlyEstablished = new string[] { "ON ", "OFF" };
                string[] ContractAutomaticRenewalCategory = new string[] { "True ", "False " };
                string[] RenewableEnergy = new string[] { "None ", "Exemption " };
                string[] EnvironmentCategory = new string[] { "True ", " False" };
                string[] SelectedWhenSupply = new string[] { "Load following ", "Base " };
                string[] Sparelinetargetclassification = new string[] { "True ", "False " };
                string[] Standbypowersupplytargetcategory = new string[] { "True ", " False" };
                string[] Selfsupportingtargetclassification = new string[] { "True ", " False" };


                for (int i = 0; i < contractsData.Length; i++)
                {
                    sheet2.Cells[i + 2, 4].Value = contractsData[i];
                }

                for (int i = 0; i < CurrentSupplyClass.Length; i++)
                {
                    sheet2.Cells[i + 2, 5].Value = CurrentSupplyClass[i];
                }

                for (int i = 0; i < Initiative.Length; i++)
                {
                    sheet2.Cells[i + 2, 6].Value = Initiative[i];
                }

                for (int i = 0; i < PowerConfiguration.Length; i++)
                {
                    sheet2.Cells[i + 2, 7].Value = PowerConfiguration[i];
                }

                for (int i = 0; i < CertificateUsage.Length; i++)
                {
                    sheet2.Cells[i + 2, 8].Value = CertificateUsage[i];
                }

                for (int i = 0; i < Mainsupplycategory.Length; i++)
                {
                    sheet2.Cells[i + 2, 10].Value = Mainsupplycategory[i];
                }

                for (int i = 0; i < CustomerTypes.Length; i++)
                {
                    sheet2.Cells[i + 2, 11].Value = CustomerTypes[i];
                }

                for (int i = 0; i < BillingMethod.Length; i++)
                {
                    sheet2.Cells[i + 2, 12].Value = BillingMethod[i];
                }

                for (int i = 0; i < PublicOfficeprivateSector.Length; i++)
                {
                    sheet2.Cells[i + 2, 15].Value = PublicOfficeprivateSector[i];
                }

                for (int i = 0; i < NewContinuationCategory.Length; i++)
                {
                    sheet2.Cells[i + 2, 16].Value = NewContinuationCategory[i];
                }

                for (int i = 0; i < MainFeeStructures.Length; i++)
                {
                    sheet2.Cells[i + 2, 17].Value = MainFeeStructures[i];
                }

                for (int i = 0; i < Estimate.Length; i++)
                {
                    sheet2.Cells[i + 2, 18].Value = Estimate[i];
                }

                for (int i = 0; i < NewlyEstablished.Length; i++)
                {
                    sheet2.Cells[i + 2, 19].Value = NewlyEstablished[i];
                }

                for (int i = 0; i < ContractAutomaticRenewalCategory.Length; i++)
                {
                    sheet2.Cells[i + 2, 20].Value = ContractAutomaticRenewalCategory[i];
                }

                for (int i = 0; i < RenewableEnergy.Length; i++)
                {
                    sheet2.Cells[i + 2, 21].Value = RenewableEnergy[i];
                }

                for (int i = 0; i < EnvironmentCategory.Length; i++)
                {
                    sheet2.Cells[i + 2, 22].Value = EnvironmentCategory[i];
                }

                for (int i = 0; i < SelectedWhenSupply.Length; i++)
                {
                    sheet2.Cells[i + 2, 23].Value = SelectedWhenSupply[i];
                }

                for (int i = 0; i < Sparelinetargetclassification.Length; i++)
                {
                    sheet2.Cells[i + 2, 24].Value = Sparelinetargetclassification[i];
                }

                for (int i = 0; i < Standbypowersupplytargetcategory.Length; i++)
                {
                    sheet2.Cells[i + 2, 25].Value = Standbypowersupplytargetcategory[i];
                }

                for (int i = 0; i < Selfsupportingtargetclassification.Length; i++)
                {
                    sheet2.Cells[i + 2, 26].Value = Selfsupportingtargetclassification[i];
                }

                FileInfo fileInfo = new FileInfo("D:\\Csharpwork\\BatchEnrollment.xlsx");
                package.SaveAs(fileInfo);
                return "Excel Create";
            }

        }
    }
}   
