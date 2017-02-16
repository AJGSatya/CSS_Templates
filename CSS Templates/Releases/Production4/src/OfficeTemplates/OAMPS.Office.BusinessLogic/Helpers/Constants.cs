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
            public const string CurrentInsuerData = @"currentInsurerData";
            public const string ReccomendedInsuerData = @"reccomendedInsurerData";
            public const string PreRenewalQuestionareQuestions = @"factfinderQuestions";
            public const string QuoteSlipSchedules = @"QuoteSlipSchedules";
            public const string NoWizard = @"NoWizard";
            public const string GenerateQuoteSlip = @"GenerateQuoteSlip";
        }

        public static class Configuration
        {
            public const string VersionNumber = @"2.0"; //also update the installer package
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
            public const string ABNKey = @"ABNKey";
            public const string AFSLKey = @"AFSLKey";
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
            public const string ABN = @"ABN";
            public const string AFSL = @"AFSL";
            public const string Category = @"Category";
            public const string Website = @"Website";
            public const string Key = @"Key";
            public const string FactFinder = @"FactFinder";
            public const string QuoteSlipSchedulesFfLookupId = @"Questionares";

            public const string PreRenewalQuestionareMappingsPolicyClasses = @"PolicyClasses";
        }

        public static class SharePointQueries
        {
            //" + templateName + "
            //" + heading + "
            public const string HelpGetItemQuery = @"<View>" +
                                                   "<Query>" +
                                                   "<Where>" +
                                                   "<And>" +
                                                   "<Eq><FieldRef Name='Template' /><Value Type='Lookup'>{0}</Value></Eq>" +
                                                   "<Eq><FieldRef Name='Title' /><Value Type='Text'>{1}</Value></Eq>" +
                                                   "</And>" +
                                                   "</Where>" +
                                                   "</Query>" +
                                                   "</View>";

            public const string GetItemByTitleQuery = @"<View>" +
                                                      "<Query>" +
                                                      "<Where>" +
                                                      "<Eq><FieldRef Name='Title' /><Value Type='Text'>{0}</Value></Eq>" +
                                                      "</Where>" +
                                                      "</Query>" +
                                                      "</View>";


            public const string GetItemByQuestionaresLookupId = @"<View>" +
                                                 "<Query>" +
                                                 "<Where>" +
                                                 "<Eq><FieldRef Name='Questionares' /><Value Type='Lookup'>{0}</Value></Eq>" +
                                                 "</Where>" +
                                                 "</Query>" +
                                                 "</View>";


            public const string GetItemByPolicyTypeQuery = @"<View>" +
                                                           "<Query>" +
                                                           "<Where>" +
                                                           "<Eq><FieldRef Name='Policy_x0020_Type' /><Value Type='Lookup'>{0}</Value></Eq>" +
                                                           "</Where>" +
                                                           "</Query>" +
                                                           "</View>";

            public const string AllItemsSortBySortOrder =
                @"<View><Query><OrderBy><FieldRef Name='SortOrder'/></OrderBy></Query></View>";

            public const string AllItemsSortByTitle =
                @"<View><Query><OrderBy><FieldRef Name='Title'/></OrderBy></Query></View>";

            public const string RenewalLetterFragmentsByKey = @"<View><Query>
                                                               <Where>
                                                                  <Or>
                                                                     <Eq>
                                                                        <FieldRef Name='Key' />
                                                                        <Value Type='Text'>General Advice Warning</Value>
                                                                     </Eq>
                                                                     <Or>
                                                                        <Eq>
                                                                           <FieldRef Name='Key' />
                                                                           <Value Type='Text'>Uninsured Risks Checklist</Value>
                                                                        </Eq>
                                                                        <Or>
                                                                           <Eq>
                                                                              <FieldRef Name='Key' />
                                                                              <Value Type='Text'>Privacy Statement</Value>
                                                                           </Eq>
                                                                           <Or>
                                                                              <Eq>
                                                                                 <FieldRef Name='Key' />
                                                                                 <Value Type='Text'>Important Notices</Value>
                                                                              </Eq>
                                                                              <Eq>
                                                                                 <FieldRef Name='Key' />
                                                                                 <Value Type='Text'>Financial Services Guide letter</Value>
                                                                              </Eq>
                                                                           </Or>
                                                                        </Or>
                                                                     </Or>
                                                                  </Or>
                                                               </Where>
                                                            </Query></View>";


            public const string PolicyClassesWithFragemnts2 =
                @"<View>  
                    <Query/>

                       <ProjectedFields>
  <Field 
    Name='Policy_x0020_Type_x003a_ID'
    Type='Lookup'
    List='CamlTest'
    ShowField='Policy_x0020_Type_x003a_ID'/>
</ProjectedFields>
                    <Query>                                      
                    </Query>
                    <Joins>
                      <Join Type='LEFT' ListAlias='CamlTest'>
                        <Eq>
                          <FieldRef Name='Policy_x0020_Type_x003a_ID' RefType='Id'/>
                          <FieldRef List='CamlTest' Name='ID'/>
                        </Eq>
                      </Join>
                    </Joins>
                    <ProjectedFields>
                      <Field
                        Name='OrderingCustomer'
                        Type='Lookup' 
                        List='CamlTest'                       
                        ShowField='Title' />
                    </ProjectedFields>


                </View>";


            public const string PolicyClassesWithFragemnts =
                @"<View><Query>
                    <Joins>
                      <Join Type='LEFT' ListAlias='PolicyFragments'>
                        <Eq>
                          <FieldRef Name='Policy_x0020_Type_x003a_ID' RefType='Id'/>
                          <FieldRef List='PolicyFragments' Name='ID'/>
                        </Eq>
                      </Join>
                    </Joins>
                    <ProjectedFields>
                      <Field
                        Name='FragmentTitle'
                        Type='Lookup'
                        List='PolicyFragments'
                        ShowField='Policy_x0020_Type' />
                    </ProjectedFields>

                </Query></View>";

            public static string LogDurationGetQuery = @"<View>" +
                                                       "<Query>" +
                                                       "<Where>" +
                                                       "<Eq><FieldRef Name='DocumentID' /><Value Type='Text'>{0}</Value></Eq>" +
                                                       "</And>" +
                                                       "</Where>" +
                                                       "</Query>" +
                                                       "</View>";
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
            public const string RenewalLetterAttachments = @"RenewalLetterAttachments";
            public const string ApprovalForm = @"ApprovalForm";
            public const string ClaimsMadeWarning = @"ClaimsMadeWarning";
            public const string BasisOfCoverPrevious = @"BasisOfCoverPrevious";
            public const string PreviousClaimsProcedure = @"PreviousClaimsProcedure";
            public const string FactFinderStart = @"FactFinderStart";
            public const string FactFinderEnd = @"FactFinderEnd";


        }

        public static class WordContentControls
        {
            public const string QuoteSlipTitle = "QuoteSlipTitle";
            public const string DocumentTitle = "DocumentTitle";
            public const string Instructions = "Instructions";
        }

        public static class WordDocumentProperties
        {
            public const string BuiltInTitle = @"Title";
            public const string CoverPageTitle = @"Cover Page Title";
            public const string UFI = @"UFI";
            public const string ApprovalForm = @"ApprovalForm";
            public const string ClaimsMadeWarning = @"ClaimsMadeWarning";
            public const string ClientProfile = @"ClientProfileProp";
            public const string LogoTitle = @"Logo Title";
            public const string IncludedPolicyTypes = @"Included Policy Types";
            public const string Segment = @"segment";
            public const string RenewalPolicy = @"PolicyName";
            public const string StatutoryInformation = @"statutory information";
            public const string Remuneration = @"Remuneration";
            public const string CacheName = @"CacheName";
            public const string IncludedPolicyTypesCount = @"Included Policy Types Count";
            public const string ActiveDocumentDuration = @"1337";


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
        }

        public static class WordTables
        {
            public const string RenewalLetterPolicies = @"Policies";
            public const string RenewalReportPremiumCosts = @"Premium Costs";
            public const string RenewalReportPremiumSummary = @"Premium Summary";
            public const string InsuranceManualProgramSummary = @"Insurance Program Summary";
        }
    }
}