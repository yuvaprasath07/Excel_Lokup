using ClosedXML.Excel;
using Entity;
using Microsoft.AspNetCore.Http;


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

            using (var workbook = new XLWorkbook())
            {
                var Sheet1 = workbook.Worksheets.Add("BatchEnrollment");
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
                    Sheet1.Range("A1:GH1").Style.Font.Bold = true;
                    Sheet1.Range("A1:GH1").Style.Fill.BackgroundColor = XLColor.DarkGray;
                    Sheet1.Cell(1, i + 1).Value = columnHeaders[i];
                }

                var Sheet2 = workbook.Worksheets.Add("Lookup");

                Sheet2.Cell("A1").Value = "Self-supplementary power calculation method*";
                Sheet2.Cell("B1").Value = "Payment Method*";
                Sheet2.Cell("C1").Value = "Load pattern";
                Sheet2.Cell("D1").Value = "Contracts";
                Sheet2.Cell("E1").Value = "Current Supply Class";
                Sheet2.Cell("F1").Value = "Initiative";
                Sheet2.Cell("G1").Value = "Power configuration";
                Sheet2.Cell("H1").Value = "Certificate usage";
                Sheet2.Cell("I1").Value = "PowerSupplyContractTypes(contracts2)";
                Sheet2.Cell("J1").Value = "Main_supply category";
                Sheet2.Cell("K1").Value = "Customer Types";
                Sheet2.Cell("L1").Value = "Billing Method";
                Sheet2.Cell("M1").Value = "Utility";
                Sheet2.Cell("N1").Value = "Area";
                Sheet2.Cell("O1").Value = "Public office/private sector";
                Sheet2.Cell("P1").Value = "New continuation category";
                Sheet2.Cell("Q1").Value = "Main fee structure";
                Sheet2.Cell("R1").Value = "Estimate";
                Sheet2.Cell("S1").Value = "Newly established";
                Sheet2.Cell("T1").Value = "Contract automatic renewal category";
                Sheet2.Cell("U1").Value = "Renewable energy surcharge exemption target category";
                Sheet2.Cell("V1").Value = "Environment menu target category";
                Sheet2.Cell("W1").Value = "Selected when main_partial supply";
                Sheet2.Cell("X1").Value = "Spare line target classification";
                Sheet2.Cell("Y1").Value = "Standby power supply target category";
                Sheet2.Cell("Z1").Value = "Self-supporting target classification";
                Sheet2.Range("A1:Z1").Style.Font.Bold = true;
                Sheet2.Range("A1:Z1").Style.Fill.BackgroundColor = XLColor.DarkGray;

                var SelfSupplementaryPowerCalculationMethod = _dbContext.L_SelfSupplementaryPowerCalculationMethod.ToList();
                var paymentmethod = _dbContext.L_PaymentMethod.ToList();
                var Loadpattern = _dbContext.L_Loadpattern.ToList();
                var PowerSupplyContractTypesContracts2 = _dbContext.L_PowerSupplyContractTypesContracts2.ToList();
                var utility = _dbContext.L_Utility.ToList();
                var area = _dbContext.L_Area.ToList();
                var MainFeeStructure = _dbContext.L_MainFeeStructure.ToList();

                for (int row = 2; row <= SelfSupplementaryPowerCalculationMethod.Count + 1; row++)
                {
                    var currentItem = SelfSupplementaryPowerCalculationMethod[row - 2];
                    Sheet2.Cell("A" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= paymentmethod.Count + 1; row++)
                {
                    var currentItem = paymentmethod[row - 2];
                    Sheet2.Cell("B" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= Loadpattern.Count + 1; row++)
                {
                    var currentItem = Loadpattern[row - 2];
                    Sheet2.Cell("C" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= PowerSupplyContractTypesContracts2.Count + 1; row++)
                {
                    var currentItem = PowerSupplyContractTypesContracts2[row - 2];
                    Sheet2.Cell("I" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= utility.Count + 1; row++)
                {
                    var currentItem = utility[row - 2];
                    Sheet2.Cell("M" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= area.Count + 1; row++)
                {
                    var currentItem = area[row - 2];
                    Sheet2.Cell("N" + row).Value = currentItem.Description;
                }
                for (int row = 2; row <= MainFeeStructure.Count + 1; row++)
                {
                    var currentItem = MainFeeStructure[row - 2];
                    Sheet2.Cell("Q" + row).Value = currentItem.Description;
                }


                string[] contract = new string[] { "Actual demand", "Consultation" };
                string[] currentsupply = new string[] { "Full supply", "Partial supply" };
                string[] initiative = new string[] { "CDP", "SBT", "RE100" };
                string[] powerconfiguration = new string[] { "With specific request", "Without specific request" };
                string[] certifiacteusage = new string[] { "With specific request", "Without specific request" };
                string[] supplycategory = new string[] { "Full supply", "Partial supply" };
                string[] customertypes = new string[] { "Residential", "Commercial" };
                string[] billingmethod = new string[] { "Web", "Mail" };
                string[] publicprivate = new string[] { "Public", "Private" };
                string[] newcatagory = new string[] { "New", "Continuous" };
                string[] estimate = new string[] { "New", "Add base", "Change unit price" };
                string[] newlyestablished = new string[] { "ON", "OFF" };
                string[] renewalcategory = new string[] { "True", "False" };
                string[] renewableenergy = new string[] { "None", "Exemption" };
                string[] environmentmenu = new string[] { "True", "False" };
                string[] partialsupply = new string[] { "Load following", "Base" };
                string[] spareline = new string[] { "True", "False" };
                string[] standby = new string[] { "True", "False" };
                string[] selfsupporting = new string[] { "True", "False" };

                for (int row = 2; row <= contract.Length + 1; row++)
                {
                    var currentItem = contract[row - 2];
                    Sheet2.Cell("D" + row).Value = currentItem;
                }
                for (int row = 2; row <= currentsupply.Length + 1; row++)
                {
                    var currentItem = currentsupply[row - 2];
                    Sheet2.Cell("E" + row).Value = currentItem;
                }
                for (int row = 2; row <= initiative.Length + 1; row++)
                {
                    var currentItem = initiative[row - 2];
                    Sheet2.Cell("F" + row).Value = currentItem;
                }
                for (int row = 2; row <= powerconfiguration.Length + 1; row++)
                {
                    var currentItem = powerconfiguration[row - 2];
                    Sheet2.Cell("G" + row).Value = currentItem;
                }
                for (int row = 2; row <= certifiacteusage.Length + 1; row++)
                {
                    var currentItem = certifiacteusage[row - 2];
                    Sheet2.Cell("H" + row).Value = currentItem;
                }
                for (int row = 2; row <= supplycategory.Length + 1; row++)
                {
                    var currentItem = supplycategory[row - 2];
                    Sheet2.Cell("j" + row).Value = currentItem;
                }
                for (int row = 2; row <= customertypes.Length + 1; row++)
                {
                    var currentItem = customertypes[row - 2];
                    Sheet2.Cell("k" + row).Value = currentItem;
                }
                for (int row = 2; row <= billingmethod.Length + 1; row++)
                {
                    var currentItem = billingmethod[row - 2];
                    Sheet2.Cell("L" + row).Value = currentItem;
                }
                for (int row = 2; row <= publicprivate.Length + 1; row++)
                {
                    var currentItem = publicprivate[row - 2];
                    Sheet2.Cell("O" + row).Value = currentItem;
                }
                for (int row = 2; row <= newcatagory.Length + 1; row++)
                {
                    var currentItem = newcatagory[row - 2];
                    Sheet2.Cell("P" + row).Value = currentItem;
                }
                for (int row = 2; row <= estimate.Length + 1; row++)
                {
                    var currentItem = estimate[row - 2];
                    Sheet2.Cell("R" + row).Value = currentItem;
                }
                for (int row = 2; row <= newlyestablished.Length + 1; row++)
                {
                    var currentItem = newlyestablished[row - 2];
                    Sheet2.Cell("S" + row).Value = currentItem;
                }
                for (int row = 2; row <= renewalcategory.Length + 1; row++)
                {
                    var currentItem = renewalcategory[row - 2];
                    Sheet2.Cell("T" + row).Value = currentItem;
                }
                for (int row = 2; row <= renewableenergy.Length + 1; row++)
                {
                    var currentItem = renewableenergy[row - 2];
                    Sheet2.Cell("U" + row).Value = currentItem;
                }
                for (int row = 2; row <= environmentmenu.Length + 1; row++)
                {
                    var currentItem = environmentmenu[row - 2];
                    Sheet2.Cell("V" + row).Value = currentItem;
                }
                for (int row = 2; row <= partialsupply.Length + 1; row++)
                {
                    var currentItem = partialsupply[row - 2];
                    Sheet2.Cell("W" + row).Value = currentItem;
                }
                for (int row = 2; row <= spareline.Length + 1; row++)
                {
                    var currentItem = spareline[row - 2];
                    Sheet2.Cell("X" + row).Value = currentItem;
                }
                for (int row = 2; row <= standby.Length + 1; row++)
                {
                    var currentItem = standby[row - 2];
                    Sheet2.Cell("Y" + row).Value = currentItem;
                }
                for (int row = 2; row <= selfsupporting.Length + 1; row++)
                {
                    var currentItem = selfsupporting[row - 2];
                    Sheet2.Cell("Z" + row).Value = currentItem;
                }


                //LokkupData
                var calculationRange = Sheet2.Range(2, 1, SelfSupplementaryPowerCalculationMethod.Count + 1, 1);
                var customerTypeValidation = Sheet1.Range("FW2:FW1000").SetDataValidation();
                customerTypeValidation.List(calculationRange);

                var paymentmethodRange = Sheet2.Range(2, 2, paymentmethod.Count + 1, 2);
                var paymentmethodValidation = Sheet1.Range("BI2:BI1000").SetDataValidation();
                paymentmethodValidation.List(paymentmethodRange);

                var LoadpatternRange = Sheet2.Range(2, 3, Loadpattern.Count + 1, 3);
                var loadValidation = Sheet1.Range("CH2:CH1000").SetDataValidation();
                loadValidation.List(LoadpatternRange);

                var utilityRange = Sheet2.Range(2, 13, utility.Count + 1, 13);
                var utilitys = Sheet1.Range("CJ1:CJ1000").SetDataValidation();
                utilitys.List(utilityRange);

                var areaRange = Sheet2.Range(2, 13, area.Count + 1, 13);
                var areas = Sheet1.Range("BY1:BY1000").SetDataValidation();
                areas.List(areaRange);

                var MainFeeStructureRange = Sheet2.Range(2, 17, MainFeeStructure.Count + 1, 17);
                var MainFeeStructures = Sheet1.Range("CG1:CG1000").SetDataValidation();
                MainFeeStructures.List(MainFeeStructureRange);

                var publicprivateRange = Sheet2.Range(2, 15, publicprivate.Length + 1, 15);
                var publicprivates = Sheet1.Range("C1:C11000").SetDataValidation();
                publicprivates.List(publicprivateRange);

                var customertypesRange = Sheet2.Range(2, 11, customertypes.Length + 1, 11);
                var customertype = Sheet1.Range("D1:D11000").SetDataValidation();
                customertype.List(customertypesRange);

                var billingmethodRange = Sheet2.Range(2, 12, billingmethod.Length + 1, 12);
                var billingmethods = Sheet1.Range("BJ1:BJ11000").SetDataValidation();
                billingmethods.List(billingmethodRange);

                var newlyestablishedRange = Sheet2.Range(2, 18, newlyestablished.Length + 1, 18);
                var newlyestablisheds = Sheet1.Range("BK1:BK11000").SetDataValidation();
                newlyestablisheds.List(newlyestablishedRange);

                var newcatagoryRange = Sheet2.Range(2, 16, newcatagory.Length + 1, 16);
                var newcatagorys = Sheet1.Range("BL1:BL11000").SetDataValidation();
                newcatagorys.List(newcatagoryRange);

                var contractRange = Sheet2.Range(2, 4, contract.Length + 1, 4);
                var contracts = Sheet1.Range("CF1:CF11000").SetDataValidation();
                contracts.List(contractRange);

                var renewalcategoryRange = Sheet2.Range(2, 20, renewalcategory.Length + 1, 20);
                var renewalcategorys = Sheet1.Range("CN1:CN11000").SetDataValidation();
                renewalcategorys.List(renewalcategoryRange);

                var supplycategoryRange = Sheet2.Range(2, 5, supplycategory.Length + 1, 5);
                var supplycategorys = Sheet1.Range("CU1:CU11000").SetDataValidation();
                supplycategorys.List(supplycategoryRange);

                var renewableenergyRange = Sheet2.Range(2, 21, renewableenergy.Length + 1, 21);
                var renewableenergys = Sheet1.Range("DY1:DY11000").SetDataValidation();
                renewableenergys.List(renewableenergyRange);

                var environmentmenuRange = Sheet2.Range(2, 22, environmentmenu.Length + 1, 22);
                var environmentmenus = Sheet1.Range("EC1:EC11000").SetDataValidation();
                environmentmenus.List(environmentmenuRange);

                var initiativeRange = Sheet2.Range(2, 6, initiative.Length + 1, 6);
                var initiatives = Sheet1.Range("ED1:ED11000").SetDataValidation();
                initiatives.List(initiativeRange);


                var certifiacteusageRange = Sheet2.Range(2, 7, certifiacteusage.Length + 1, 7);
                var certifiacteusages = Sheet1.Range("EF1:EF11000").SetDataValidation();
                certifiacteusages.List(certifiacteusageRange);

                var PowerSupplyContractTypesContracts2Range = Sheet2.Range(2, 9, PowerSupplyContractTypesContracts2.Count + 1, 9);
                var PowerSupplyContractTypesContracts2s = Sheet1.Range("EI1:EI11000").SetDataValidation();
                PowerSupplyContractTypesContracts2s.List(PowerSupplyContractTypesContracts2Range);

                var currentsupplyRange = Sheet2.Range(2, 10, currentsupply.Length + 1, 10);
                var currentsupplys = Sheet1.Range("EN1:EN11000").SetDataValidation();
                currentsupplys.List(currentsupplyRange);

                var partialsupplyRange = Sheet2.Range(2, 23, partialsupply.Length + 1, 23);
                var partialsupplys = Sheet1.Range("EO1:EO11000").SetDataValidation();
                partialsupplys.List(partialsupplyRange);

                var sparelineRange = Sheet2.Range(2, 24, spareline.Length + 1, 24);
                var sparelines = Sheet1.Range("FJ1:FJ11000").SetDataValidation();
                sparelines.List(sparelineRange);

                var standbyRange = Sheet2.Range(2, 25, standby.Length + 1, 25);
                var standbys = Sheet1.Range("FO1:FO11000").SetDataValidation();
                standbys.List(standbyRange);

                var selfsupportingRange = Sheet2.Range(2, 26, selfsupporting.Length + 1, 26);
                var selfsupportings = Sheet1.Range("FT1:FT11000").SetDataValidation();
                selfsupportings.List(selfsupportingRange);

                workbook.SaveAs("D:\\Csharpwork\\ZN-BatchEnrollment.xlsx");
                return "Success";
            }


        }

        public object UploadFileAsync(IFormFile file)
        {
            string filePath = @"D:\Csharpwork\Filupload.xlsx";
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
            try
            {
                if (file == null || file.Length == 0)
                {
                    return "No file uploaded.";
                }

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                using (var uploadedWorkbook = new XLWorkbook(filePath))
                {
                    var worksheet = uploadedWorkbook.Worksheets.FirstOrDefault();
                    var uploadedHeaders = worksheet?.Row(1).CellsUsed().Select(cell => cell.Value.ToString().ToLower()).ToList();


                    List<string> mismatchedHeaders = new List<string>();

                    if (uploadedHeaders == null || uploadedHeaders.Count != columnHeaders.Length)
                    {
                        mismatchedHeaders.Add("Mismatched header count.");
                    }
                    else
                    {
                        for (int i = 0; i < columnHeaders.Length; i++)
                        {
                            if (!uploadedHeaders[i].Equals(columnHeaders[i].ToLower()))
                            {
                                mismatchedHeaders.Add($"{i + 1}. {columnHeaders[i]}");
                            }
                        }
                    }

                    if (mismatchedHeaders.Count > 0)
                    {
                        File.Delete(filePath);
                        return mismatchedHeaders;
                    }

                    return $"File uploaded to: {filePath}";
                }
            }
            catch (Exception ex)
            {
                return $"Internal server error: {ex.Message}";
            }
        }
    }
}


