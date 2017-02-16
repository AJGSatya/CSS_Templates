namespace OAMPS.Office.BusinessLogic.Helpers
{
    public static class Constants
    {
        public static class CacheNames
        {
            public const string MajorPolicyClassItems = @"MajorPolicyClassItems";
            public const string MinorPolicyClassItems = @"MinorPolicyClassItems";
            public const string MajorQuestionClassItems = @"MajorQuestionClassItems";
            public const string MinorQuestionClassItems = @"MinorQuestionClassItems";
            public const string RegenerateTemplate = @"RegenerateTemplate";
            public const string RenLtrFragments = @"RenLtrFragments";
            public const string QuoteSlipFragments = @"QuoteSlipFragments";
            public const string CurrentInsuerData = @"currentInsurerData";
            public const string ReccomendedInsuerData = @"reccomendedInsurerData";
            public const string PreRenewalQuestionareQuestions = @"factfinderQuestions";
            public const string QuoteSlipSchedules = @"QuoteSlipSchedules";
            public const string NoWizard = @"NoWizard";
            public const string ConvertWizard = @"ConvertWizard";
            public const string GenerateQuoteSlip = @"GenerateQuoteSlip";
            public const string ServicePlanImages = @"ServicePlanImages";
            public const string ShortFormProposalDocumentTitles = @"ShortFormProposalDocumentTitles";
            public const string NoCache = @"No Cache";
        }

        public static class Configuration
        {
            public const string VersionNumber = @"3.6"; //also update the installer packages x2
        }

        public static class ControlNames
        {
            public const string TabPageCoverPagesPictureBoxName = "pictureBoxImage";
            public const string TabPageCoverPagesName = @"Themes";
            public const string TabPageCoverPagesTitle = @"Themes";
            public const string TabPageLogosName = @"Logos";
            public const string TabPageLogosTitle = @"Logos";
            public const string TabControl = @"tbcWizardScreens";

            public const string TxtClientName = @"txtClientName";
            public const string TxtExecutiveDepartment = @"txtExecutiveDepartment";
            public const string TxtExecutiveName = @"txtExecutiveName";
        }

        public static class FragmentKeys
        {
            public const string GeneralAdviceWarning = @"General Advice Warning";
            public const string UninsuredRisksReviewList = @"Uninsured Risks Checklist";
            public const string PrivacyStatement = @"Privacy Statement";
            public const string StatutoryNotices = @"Important Notices";
            public const string FinancialServicesGuide = @"Financial Services Guide";
            public const string FinancialServicesGuideLetter = @"Financial Services Guide letter";
        }

        public static class ImageProperties
        {
            public const string Theme = @"Themes";
            public const string CompanyLogo = @"Logo";
        }

        public static class Miscellaneous
        {
            public const string ConfirmMsg =
                @"You are now changing significant document contents, a NEW document will generate for your use.";

            public const string RegenerateOnLoadMsg =
                @"Your new document is ready, please recheck your data before clicking finish.";
        }

        public static class RadioButtonValues
        {
            public const string ImageUrl = @"ImageUrl";
            public const string AbnKey = @"ABNKey";
            public const string AfslKey = @"AFSLKey";
            public const string WebsiteKey = @"WebSite";
            public const string OAMPSCompanyName = @"OAMPSCompanyName";
            public const string LongBrandingDesc = @"LongBrandingDesc";
        }

        public static class Seperators
        {
            public const string Lineseperator = "#-n";
            public const string Spaceseperator = "#-s";
        }

        public static class SharePointFields
        {
            public const string FileRef = @"FileRef";
            public const string Title = @"Title";
            public const string HeaderType = @"HeaderType";
            public const string SortOrder = @"SortOrder";
            public const string MajorPolicyClass = @"Class_x0020_of_x0020_Insurance";
            public const string TopQuestionClass = @"Category";
            public const string SubQuestionClass = @"Question_x0020_Group";
            public const string FieldId = @"ID";
            public const string LongDescription = @"Long_x0020_Description";
            public const string ShortDescription = @"Short_x0020_Description";
            public const string Content = @"Content";
            public const string WizardHelp = @"Wizard Help";

            public const string MinorClass = @"Policy_x0020_Type";
            public const string Abn = @"ABN";
            public const string Afsl = @"AFSL";
            public const string Category = @"Category";
            public const string Website = @"Website";
            public const string Key = @"Key";
            public const string FactFinder = @"FactFinder";
            public const string QuoteSlipSchedulesFfLookupId = @"FactFinder";

            public const string PreRenewalQuestionareMappingsPolicyClasses = @"PolicyClasses";

            public const string Speciality = @"Speciality";
        }

        public static class Spotlight
        {
            public const string SpotlightParentKeyUrl = @"Software\Microsoft\Office\14.0\Common";
            public const string SpotlightKeyUrl = @"Software\Microsoft\Office\14.0\Common\Spotlight";
            public const string ProviderKeyUrl = @"Software\Microsoft\Office\14.0\Common\Spotlight\Providers";
            public const string ContentKeyUrl = @"Software\Microsoft\Office\14.0\Common\Spotlight\Content";
        }

        public static class TemplateNames
        {
            public const string PlacementSlip = @"Placement Slip";
            public const string InsuranceRenewalReport = @"Insurance Renewal Report";
            public const string PreRenewalAgenda = @"Meeting Agenda";
            public const string FileNote = @"File Note";
            public const string RenewalLetter = @"Renewal Letter";
            public const string ClientDiscovery = @"Discovery Guide";
            public const string PreRenewalQuestionnaire = @"Pre-Renewal Questionnaire";
            public const string GenericLetter = @"Generic Letter";
            public const string QuoteSlip = @"Quote Slip";
            public const string InsuranceManual = @"Insurance Manual";
            public const string ShortFormProposal = @"Insurance proposal - short form";
            public const string BlankFragmentBaseOlly = @"Blank Fragment Base Olly";
        }

        public static class WordBookmarks
        {
            public const string Renumeration = @"Remuneration";
            public const string AddRenumeration = @"AddRemuneration";
            public const string RenewalLetterMain = @"RenewalLetterMain";
            public const string RenewalLetterSub = @"RenewalLetterSub";
            public const string ImportantNotes = @"ImportantNotes";
            public const string ClientDiscoveryQuestions = @"ClientDiscoveryQuestions";
            public const string ExecutiveSummary = @"ExecutiveSummary";
            public const string PurposeOfReport = @"PurposeOfReport";
            public const string PurposeOfReportGeneric = @"PurposeOfReportGeneric";
            public const string AgendaMinutes = @"MinutesBookmark";
            public const string TopLine = @"TopLine";
            public const string UFIBookmark = @"UFIBookmark";
            public const string InsertClientProfile = @"ClientProfile";
            public const string InsertClientProfileEnd = @"ClientProfileEnd";
            public const string InsertContractingProcedure = @"ContractingProcedure";
            public const string ClaimsProcedures = @"ClaimsProcedures";
            public const string RenewalLetterAttachments = @"RenewalLetterAttachments";
            public const string ApprovalForm = @"ApprovalForm";
            public const string ClaimsMadeWarning = @"ClaimsMadeWarning";
            public const string BasisOfCoverPrevious = @"BasisOfCoverPrevious";
            public const string PreviousClaimsProcedure = @"PreviousClaimsProcedure";
            public const string FactFinderStart = @"FactFinderStart";
            public const string FactFinderEnd = @"FactFinderEnd";
            public const string ServiceTeam = @"ServiceTeam";
            public const string SummaryOfCoverStart = @"SummaryOfCoverStart";
            public const string SummaryOfCoverEnd = @"SummaryOfCoverEnd";
            public const string QuotationDetailsStart = @"QuotationDetailsStart";
            public const string QuotationDetailsEnd = @"QuotationDetailsEnd";
            public const string UnderwritingStart = @"UnderwritingStart";
            public const string UnderwritingEnd = @"UnderwritingEnd";
        }

        public static class WordContentControls
        {
            public const string QuoteSlipTitle = "QuoteSlipTitle";
            public const string DocumentTitle = "DocumentTitle";
            public const string Instructions = "Instructions";
            public const string ClassOfInsuranceTitle = @"ClassOfInsuranceTitle";
        }

        public static class WordDocumentProperties
        {
            public const string BuiltInTitle = @"Title";
            public const string CoverPageTitle = @"Cover Page Title";
            public const string UFI = @"UFI";
            public const string ApprovalForm = @"ApprovalForm";
            public const string ClaimsMadeWarning = @"ClaimsMadeWarning";
            public const string ClientProfile = @"ClientProfileProp";
            public const string ContractingProcedure = @"ContractingProcedureProp";
            public const string LogoTitle = @"Logo Title";
            public const string IncludedPolicyTypes = @"Included Policy Types";
            public const string Segment = @"segment";
            public const string RenewalPolicy = @"PolicyName";
            public const string StatutoryInformation = @"statutory information";
            public const string Remuneration = @"Remuneration";
            public const string CacheName = @"CacheName";
            public const string IncludedPolicyTypesCount = @"Included Policy Types Count";
            public const string ActiveDocumentDuration = @"1337";

            public const string LoadBrandingImageSelector = @"LoadBrandingImageSelector";

            public const string RlSubFragments = @"RL_subFragments";
            public const string RlChkContacted = @"RL_chkContacted";
            public const string RlChkNewClient = @"RL_chkNewClient";
            public const string RlChkFunding = @"RL_chkFunding";
            public const string RlChkGaw = @"RL_chkGAW";
            public const string RlRdoPreprint = @"RL_RdoPreprint";

            public const string DiscoveryQuestions = @"DiscoveryQuestions";

            public const string FileNoteIsClient = @"isClient";
            public const string FileNoteIsCaller = @"isCaller";
            public const string FileNoteIsUnderwriter = @"isUnderwriter";
            public const string FileNoteByPhone = @"isByPhone";
            public const string FileNoteInPerson = @"isInPerson";
            public const string FileNoteOther = @"isOther";

            public const string DocumentId = @"documentId";

            public const string ToggleLockMode = @"LockMode";

            public const string DocumentGeneratedDate = "@GeneratedDate";
            public const string UsedDateOfFragements = "@UsedDateOfFragements";
            public const string UsedDateOfTheme = "@UsedDateOfTheme";
            public const string UsedDateOfLogo = "@UsedDateOfLogo";
            public const string UsedDateOfTemplate = "@UsedDateOfTemplate";
            public const string HideTemplateCheckMessage = "@HideTemplateCheckMessage";

            public const string Speciality = "@Speciality";
        }

        public static class WordDocumentPropertyValues
        {
            public const string ToggleLockModeLocked = @"Locked";
            public const string ToggleLockModeUnlocked = @"Unlocked";
        }

        public static class WordStyles
        {
            public const string Heading1 = @"Heading 1";
            public const string Heading3Black = @"Heading 3 Black";
            public const string Bold = @"Strong";
        }

        public static class WordTables
        {
            public const string RenewalLetterPolicies = @"Policies";
            public const string RenewalReportPremiumCosts = @"Premium Costs";
            public const string RenewalReportPremiumSummary = @"Premium Summary";
            public const string InsuranceManualProgramSummary = @"Insurance Program Summary";
            public const string PolicyWording = @"Policy Wording";
            public const string AssetRiskProtection = @"Asset Risk Protection";
            public const string IncomeandotherFinancialExposures = @"Income and other Financial Exposures";
            public const string Liabilityrisksandexposures = @"Liability risks and exposures";

            public const string AccountExec = @"AccountExec";
            public const string AssistantExec = @"AssistantExec";
            public const string ClaimsExec = @"ClaimsExec";
            public const string OtherContact = @"OtherContact";
        }
    }
}