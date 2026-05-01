using ExcelMigrationTool.Models;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Text;

namespace ExcelMigrationTool.Services;

//Excel migration service

public class ExcelMigrationService : IExcelMigrationService
{
    // SQL Command timeout in seconds (10 minutes for large datasets)
    private const int SqlCommandTimeout = 600;

    // Unit ID columns that should be resolved via UnitMaster by unit name
    private static readonly HashSet<string> UnitIdColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "AmbientTemperatureUnitID",
        "TemperatureRiseDeltaTUnitID",
        "ElectricalDesignUnitID",
        "InstrumentAirPressureUnitID",
        "CWSupplyTemperatureUnitID",
        "CWSupplyPressureUnitID",
        "DesignPressureUnitID",
        "PressureDropUnitID",
        "ExhaustPressureUnitID",
        "PressureUnitID"
    };

    public static readonly Dictionary<string, string> ProjectMapping = new(StringComparer.OrdinalIgnoreCase)
        {

            { "shell_pid", "ProjectID" },
            { "record_no", "RecordNo" },
            {"shellnumber" ,"Shellnumber"},
            { "shell_last_modified_date", "UpdatedAt" },
            { "uuu_record_last_update_user", "UpdatedName" },
            { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },

            { "process_status", "ProcessStatus" },
            { "order_status", "Status" },

            { "creator_id", "CreatedName" },
            { "k__creator_id", "PrimaveraCreatedId" },

            { "uot_c_number_sdt120", "CNumber" },

            { "ucp_pm_smn", "ManagerName" },
            { "k__ci_project_manager_upk", "PrimeveraManagerID" },

            { "description", "Description" },

            { "uuu_shell_template_picker", "ProjectTemplateID" },
            { "shelllocation", "ProjectTypeMasterID" },

            { "shellname", "ProjectName" },

            { "shell_createdate", "CreatedAt" }
        };


    // Hardcoded column mapping for CommunicationProtocol table
    private static readonly Dictionary<string, string> CommunicationProtocolColumnMapping = new(StringComparer.OrdinalIgnoreCase)
{
    { "id", "CommunicationProtocolID" },
    { "record_no", "RecordNo" },
    { "project_id", "ProjectID" },

    { "status", "Status" },
    { "process_status", "ProcessStatus" },

    { "ucp_cp_250", "CommunicationProtocolFormat" },
    { "ucp_format_sdt120", "CPFormat" },
    { "ucpt_kindly_attn_sdt120", "KindlyAttn" },

    { "uot_comp_name3_sdt250", "SoldToParty" },

    { "uot_purchase_add_sdt120", "Address1" },
    { "uot_purchase_add2_sdt120", "Address2" },
    { "uot_purchase_add3_sdt120", "Address3" },

    { "uot_pur_citytxt50", "City" },
    { "uot_india_states_pd", "StateProvince" },
    { "uot_pur_countrypd", "Country" },
    { "uot_state", "StateProvinceOtherThanIndia" },

    { "uot_email1tb120", "Email" },
    { "uuu_user_fax", "Fax" },
    { "uot_phone1_sdt50", "Phone" },
    { "phone_number002", "WorkPhone" },

    // Project Leader
    { "k__uot_project_lead_dp", "ProjectLeaderPrimaveraId" },
    { "uot_project_lead_dp", "ProjectLeaderName" },
    { "email_id0", "ProjectLeaderEmail" },
    { "phone_number0", "ProjectLeaderPhoneNumber" },
    { "mobile_number0", "ProjectLeaderMobileNumber" },
    { "extention0", "ProjectLeaderExtension" },

    // Project Manager
    { "k__uot_project_co_ordinator_dp", "ProjectManagerPrimaveraId" },
    { "uot_project_co_ordinator_dp", "ProjectManagerName" },
    { "cp_email_id_dc", "ProjectManagerEmail" },
    { "cp_mobile_number_dc", "ProjectManagerMobileNumber" },
    { "cp_extension_dc", "ProjectManagerExtension" },

    // Document Controller
    { "k__cp_document_manager_dc", "DocumentControllerPrimaveraId" },
    { "cp_document_manager_dc", "DocumentControllerName" },
    { "email_id11", "DocumentControllerEmail" },
    { "phone_number11", "DocumentControllerPhoneNumber" },
    { "extention11", "DocumentControllerExtension" },
    { "mobile_number11", "DocumentControllerMobileNumber" },

    // Zonal Head
    { "k__uot_mails_to_be_sent_to_dp", "ZonalHeadPrimaveraId" },
    { "uot_mails_to_be_sent_to_dp", "ZonalHeadName" },
    { "email_id21", "ZonalHeadEmail" },
    { "phone_number21", "ZonalHeadPhoneNumber" },
    { "extention21", "ZonalHeadExtension" },
    { "mobile_number_mails", "ZonalHeadMobileNumber" },

    // IC Head
    { "k__uot_copies_to_be_sent_to_dp", "ICHeadPrimaveraId" },
    { "uot_copies_to_be_sent_to_dp", "ICHeadName" },
    { "email_id31", "ICHeadEmail" },
    { "phone_number31", "ICHeadPhoneNumber" },
    { "extention31", "ICHeadExtension" },
    { "mobile_number_copies", "ICHeadMobileNumber" },

    // HOD
    { "k__uot_head_of_department", "HODPrimaveraId" },
    { "ucp_hop_sdt120", "HODName" },
    { "email_id41", "HODEmail" },
    { "phone_number41", "HODPhoneNumber" },
    { "extention41", "HODExtension" },
    { "mobile_number_headofdept", "HODMobileNumber" },

    // Incharge
    { "k__uot_incharge_exports_dp", "InchargePrimaveraId" },
    { "uot_incharge_exports_dp", "InchargeName" },
    { "email_id51", "InchargeEmail" },
    { "phone_number51", "InchargePhoneNumber" },
    { "extention51", "InchargeExtension" },
    { "mobile_number21", "InchargeMobileNumber" },

    // HOSS
    { "k__uot_head_spares_and_service", "HOSSPrimaveraId" },
    { "uot_head_spares_and_service", "HOSSName" },
    { "email_id61", "HOSSEmail" },
    { "phone_number61", "HOSSPhoneNumber" },
    { "extention61", "HOSSExtension" },
    { "mobile_number_headspares", "HOSSMobileNumber" },

    // Misc
    { "postal_address", "PostalAddress" },
    { "ucp_boarline_sdt120", "BoardLine" },
    { "fax1", "Address_Fax" },

    // Audit fields
    { "uuu_creation_date", "CreatedAt" },
    { "k__creator_id", "PrimaveraCreatedId" },
    { "creator_id", "CreatedName" },

    { "uuu_record_last_update_date", "UpdatedAt" },
    { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
    { "uuu_record_last_update_user", "UpdatedName" },

    { "k__uot_sold_party_dp", "CustomerId" }
};

    private static readonly Dictionary<string, string> BankGuaranteeMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "BankGuaranteeID" },
            { "record_no", "RecordNo" },
            { "uuu_record_last_update_date", "UpdatedDate" },
            { "process_status", "ProcessStatus" },
            { "status", "Status" },
            { "project_id", "ProjectID" },

            { "ubg_type_of_bg_pd", "TypeOfGuarantee" },
            { "ubg_others01_sdt250", "TypeOfGuaranteeOthers" },

            { "uuu_creation_date", "CreatedAt" },
            { "k__creator_id", "PrimaveraCreatedId" },
            { "creator_id", "CreatedName" },

            { "ubg_claim_period_01dop", "ClaimPeriodDate" },
            { "ubg_date_dop", "ContractDate" },
            { "ubg_validity_dop", "ValidityDate" },

            { "ubg_contractual_order_ca", "TotalOrderValue" },
            { "ubg_equivalent_value_ca", "TotalOrderValueINR" },
            { "ubg_currency_amount_ca", "BankGuaranteeAmount" },

            { "ubg_percent_bg_da", "PercentageOfGuarantee" },
            { "ubg_exchange_rate_da", "ExchangeRate" },

            { "ubg_currency_pd01", "Currency" },

            { "ubg_contract_01pd", "IsContractCopyAttached" },
            { "ubg_draft_bank_ynpd", "IsDraftFormatAttached" },

            { "ubg_draft_gurante_format_pd", "DraftFormat" },
            { "ubg_others_sdt2501", "DraftFormatDetails" },

            { "ubg_guarantee_num_sdt250", "BankGuaranteeNo" },
            { "ubg_issuing_bank_sdt250", "IssuingBank" },

            { "ubg_gurantee_against_pd", "GuaranteeAgainst" },
            { "ubg_others_sdt2503", "GuaranteeAgainstOthers" },

            { "ubg_warranty_clause_pd", "WarrantyClause" },
            { "ubg_others2601_sdt250", "WarrantyClauseOthers" },

            { "ubg_agre_sdt_250", "ContractReferenceNo" },

            { "ubg_amrrndment_pd", "BankGuaranteeType" },

            { "ubg_remarks_ldt", "InitiatorReviewRemarks" },

            { "uot_sd", "CProjectNumber" }
    };

    // Mapping profile for OrderTransmittalLineItemBankGuarantee uploads.
    // Rows are inserted/updated in the existing BankGuarantee table.
    private static readonly Dictionary<string, string> OrderTransmittalLineItemBankGuaranteeMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "record_id", "OrderTransmittalID" },
            { "uuu_li_last_update_date", "UpdatedAt" },
            { "ubg_type_of_bg_pd", "TypeOfGuarantee" },
            { "ot_guarantee_against_pd", "GuaranteeAgainst" },
            { "ot_bg_contractual_term_sldt", "ContractualTerms" },
            { "ot_bg_guarantee_da", "PercentageOfGuarantee" },
            { "ot_bg_guarantee_amt_da", "GuaranteeAmountINR" },
            { "ot_guarantee_amt_da", "GuaranteeAmount" },
            { "uot_cur1_pd", "Currency" },
            { "ot_exchange_rate_da", "ExchangeRate" },
            { "ubg_guarantee_num_sdt250", "BankGuaranteeNo" },
            { "ot_bg_stat_pd", "BGStatus" },
            { "ot_issued_date_dop", "IssuedDate" },
            { "ot_expiry_date_dop", "ExpiryDate" },
            { "ubg_bank_name_sdt250", "IssuingBank" },
            { "project_id", "Projectid" },
            { "id", "BankGuaranteeID" }
    };

    // Hardcoded column mapping for BankGuaranteeSpares table
    private static readonly Dictionary<string, string> BankGuaranteeSparesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "BankGuaranteeSparesID" },
        { "record_no", "RecordNo" },
        { "project_id", "ProjectID" },
        { "k__uot_bp_picker_bp", "SparesOrderTransmittalID" },
        { "ubg_guarantee_num_sdt250", "BankGuaranateeNoIssued" },
        { "ubg_type_of_bg_pd", "TypeOfGuarantee" },
        { "ubg_others01_sdt250", "TypeOfGuaranteeOthers" },
        { "ubg_agre_sdt_250", "ContractReferenceNo" },
        { "ubg_date_dop", "ContractDate" },
        { "ubg_contractual_order_ca", "TotalOrderValue" },
        { "ubg_currency_pd", "Currency" },
        { "ubg_exchange_rate_da", "ExchangeRate" },
        { "ubg_equivalent_value_ca", "TotalOrderValueINR" },
        { "ubg_percent_bg_da", "PercentageOfGuarantee" },
        { "ubg_currency_amount_ca", "BankGuaranteeAmount" },
        { "ubg_warranty_clause_pd", "WarrantyClause" },
        { "ubg_others2601_sdt250", "WarrantyClauseOthers" },
        { "ubg_gurantee_against_pd", "GuaranteeAgainst" },
        { "ubg_gurantee_against_PD01", "GuaranteeAgainstOthers" },
        { "ubg_validity_dop", "ValidityDate" },
        { "ubg_claim_period_01dop", "ClaimPeriodDate" },
        { "ubg_draft_bank_ynpd", "IsDraftFormatAttached" },
        { "ubg_draft_gurante_format_pd", "DraftFormat" },
        { "ubg_others_sdt2502", "DraftFormatDetails" },
        { "ubg_remarks_ldt", "InitiatorReviewRemarksBeforeAuthorized" },
        { "ubg_contract_01pd", "IsContractCopyAttached" },
        { "ubg_init_final_remarks_ldt", "InitiatorReviewRemarks" },
        { "ubg_fin_remarks_ldt", "FinanceRemarks" },
        { "ubg_issuing_bank_sdt250", "IssuingBank" },
        { "ubg_validity_dop02", "BGValidityDuringIssual" },
        { "ubg_claim_period_03dop", "BGClaimPeriodDate" },
        { "ubg_contractual_order_ca02", "BGTotalOrderValueINR" },
        { "ubg_currency_pd01", "BGCurrency" },
        { "ubg_exchange_rate_da02", "BGExchangeRateDuringIssual" },
        { "ubg_equivalent_value_ca02", "BGEquivalentValueINR" },
        { "ubg_percent_bg_da03", "PercentageOfBankGuarantee" },
        { "ubg_currency_amount_ca02", "BankGuaranteeAmountDuringIssual" },
        { "status", "Status" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_creation_date", "CreatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "sparesprovisionalordertransmittalid", "SparesProvisionalOrderTransmittalId" },
        { "ubg_status_pd", "BGStatus" },
        { "ubg_exp_expirydate_dop", "ExpectedExpiryDate" },
        { "ubg_remarks01_ldt", "BGStatusRemarks" },
        { "spares_sales_order_number", "SalesOrderNo" }
    };

    // Hardcoded column mapping for DeviationRequest table
    private static readonly Dictionary<string, string> DeviationRequestMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "DeviationRequestId" },
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "customermasterid", "CustomerMasterId" },
        { "enduserid", "EndUserId" },
        { "k__rcca_ref_bpp", "RCCAId" },
        { "k__dev_insp_bpp", "InspectionId" },
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "dr_dr_no_smn", "DeviationRequestNo" },
        { "des_des_no_smn", "NCNo" },
        { "dev_item_desc_sdt250", "ItemDescription" },
        { "dev_item_no_sdt50", "ItemNo" },
        { "qacr_po_no_sdt250", "PurchaseOrderNo" },
        { "dev_l_no_sdt50", "LNoOrSNo" },
        { "dev_ir_no_sdt50", "InspectionReportNo" },
        { "dev_ir_date_dop", "InspectionReportDate" },
        { "qacr_component_name_sdt250", "ComponentName" },
        { "qacr_item_code_sdt250", "ItemCode" },
        { "dev_spec_req_mtb4000", "SpecifiedRequirement" },
        { "dev_dev_req_det_mtb4000", "DeviationRequestDetails" },
        { "dev_dev_req_just_mtb4000", "DeviationRequestJustification" },
        { "sqdcm_plant_pd", "Plant" },
        { "dev_quantity_ia", "Quantity" },
        { "dev_attachment1_ynpd", "Attachment" },
        { "dev_proc_req_pd", "DeviationRequestProcess" },
        { "dev_dev_type_pd", "DeviationType" },
        { "dev_type_devi_pd", "TypeOfDeviation" },
        { "dev_operation_sdt250", "Operation" },
        { "dev_mach_name_sdt250", "MachineName" },
        { "dev_cause_code_pd", "CauseCode" },
        { "vendor_select_dp", "SelectSupplier" },
        { "dev_other_pd", "Other" },
        { "dev_cat_risk_pd", "CategoryOfRisk" },
        { "dev_tech_dec_ynpd", "IsTechnicalDecisionRequired" },
        { "dev_risk_mit_meas_mtb4000", "RiskMitigationMeasures" },
        { "dev_resp_person_upk", "RiskMitigationResponsible" },
        { "dev_plan_date_dop", "RiskMitigationPlannedDate" },
        { "dev_ser_foll_mtb4000", "ServiceFollowUp" },
        { "dev_responsible2_upk", "ServiceFollowUpResponsible" },
        { "dev_plan_date2_dop", "ServiceFollowUpPlannedDate" },
        { "dev_corr_act_pd", "CorrectiveAction" },
        { "dev_other2_pd", "CorrectiveActionOther" },
        { "dev_responsible_upk", "ResponsiblePerson" },
        { "dev_doc_root_id_dop", "DateOfCompletionForIdentification" },
        { "dev_doc_root_elim_dop", "DateOfCompletionForElimination" },
        { "dev_ini_dept_pd", "InitiatorDepartment" },
        { "dev_resp_qa_upk", "ResponsibleQAEngineer" },
        { "dev_initiat_hod_name_sdt50", "HODName" },
        { "dev_initiator_hod_date_dop", "HODDate" },
        { "dev_hod_comm_sdt250", "HODComments" },
        { "dev_cause_area_pd", "CauseArea" },
        { "dev_sec_cause_pd", "SecondaryCauseArea" },
        { "dev_rec_mat_dev_pd", "RecipientOfMaterial" },
        { "dev_sec_rec_mat_dev_pd", "SecondaryRecipientOfMaterial" },
        { "dev_proj_ec_pd1", "ProjectApprovalDecision" },
        { "dev_name2_sdt50", "ProjectApprovalName" },
        { "dev_date2_dop", "ProjectApprovalDate" },
        { "dev_comments2_mtb4000", "ProjectApprovalComments" },
        { "dev_name4_sdt50", "DesignName" },
        { "dev_date4_dop", "DesignDate" },
        { "dev_remarks2_mtb4000", "DesignRemarks" },
        { "dev_design_rd_pd", "HODDesign" },
        { "dev_name_sdt50", "ReviewOfDeviationName" },
        { "dev_date_dop", "ReviewOfDeviationDate" },
        { "dev_comments_mtb4000", "RecommendationFromDesign" },
        { "dev_change_mating_ynpd", "IsChangeInMatingParts" },
        { "dev_rework_res_upk", "ReworkResponsible" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__vendor_select_dp", "VendorId" },
        { "sqdcm_rcca_required_pd", "IsRCCARequired" },
        { "sqdcm_rcca_lead_upk", "RCCALead" },
        { "rcca_lead_department", "LeadDepartment" },
        { "dev_cs_qa_pd1", "HODCSandQAReview" },
        { "dev_name3_sdt50", "HODCSandQAName" },
        { "dev_date3_dop", "HODCSandQADate" },
        { "dev_comments3_mtb4000", "HODCSandQComments" },
        { "dev_component_sdt50", "NCDrawingComponent" },
        { "dev_nc_drawing_no_sdt50", "NCDrawingNumber" },
        { "dev_item_code_sdt50", "NCDrawingItemCode" },
        { "dev_design_man_hrs_da", "NCDrawingDesignManHrs" },
        { "dev_design_remarks_mtb4000", "NCDrawingDesignRemarks" },
        { "dev_bom_updated_ynpd", "IsBOMUpdatedWithItemCodes" },
        { "dev_design_remarks2_mtb4000", "BOMDesignRemarks" },
        { "dev_design_bom_rev_ynpd", "IsDesignBOMRevised" },
        { "dev_design_recomm_mtb4000", "DesignRecommendation" },
        { "dev_details_rework_mtb4000", "DetailsOfRework" },
        { "dev_qa_observations_mtb4000", "InspectionReportQAObservation" },
        { "dev_decision_pd", "InspectionReportDecision" },
        { "dev_remarks_mtb4000", "InspectionReportRemarks" },
        { "k__dev_resp_person_upk", "PrimaveraRiskMitigationResponsibleId" },
        { "k__dev_responsible_upk", "PrimaveraResponsibleId" },
        { "k__dev_resp_qa_upk", "PrimaveraResponsibleQAEngineerId" },
        { "k__sqdcm_rcca_lead_upk", "PrimaveraRCCALeadId" },
        { "k__dev_responsible2_upk", "PrimaveraServiceFollowUpResponsibleId" },
        { "qacr_supply_type_pd", "SupplyType" },
        { "dev_ec_dec_pd", "HODECDecision" },
        { "dev_name5_sdt50", "HODECName" },
        { "dev_date5_dop", "HODECDate" },
        { "dev_comments4_mtb4000", "HODECComments" }
    };

    // Hardcoded column mapping for DesignRecommendation table
    private static readonly Dictionary<string, string> DesignRecommendationMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "DesignRecommendationId" },
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "customermasterid", "CustomerMasterId" },
        { "enduserid", "EndUserId" },
        { "k__dev_insp_bpp", "InspectionId" },
        { "dev_rework_res_upk", "ReworkResponsible" },
        { "k__dev_sel_deviation_bpp", "DeviationRequestId" },
        { "dev_design_recomm_mtb4000", "DesignRecommendation" },
        { "dev_details_rework_mtb4000", "DetailsOfRework" },
        { "dev_qa_observations_mtb4000", "QAObservations" },
        { "dev_decision_pd", "InspectionReportDecision" },
        { "dev_remarks_mtb4000", "InspectionReportRemarks" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "record_no", "RecordNo" },
        { "status", "Status" }
    };

    // Hardcoded column mapping for FrameMaster table
    private static readonly Dictionary<string, string> FrameMasterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "id", "FrameMasterID" },
        { "status", "Status" },
        { "fm_frame_sdt50", "Frame" },
        { "fm_frame_plan_sdt50", "FrameForPlan" },
        { "fm_turb_del_week_da", "TurbineDelivery" },
        { "fm_turb_del_days_da", "DispatchDays" },
        { "fm_rm_avail_patt_da", "TurbineDeliveryWithoutRMAvailable" },
        { "fm_rm_new_patt_da", "TurbineDeliveryWithoutRMNew" },
        { "fm_turb_mrt_da", "TurbineMRT" },
        { "fm_turb_assem_days_da", "TurbineAssembly" },
        { "fm_sc_casing_days_da", "SteamCasing" },
        { "fm_ex_casing_fab_days_da", "ExCasingFabrication" },
        { "fm_ex_casing_mcg_days_da", "ExCasingMcg" },
        { "fm_ex_casing_fab_mc_da", "ExhaustCasing" },
        { "fm_rotor_mcg_days_da", "Rotor" },
        { "fm_blade_mcg_days_da", "Blades" },
        { "fm_pede_fab_mcng_da", "PedestalFabrication" },
        { "fm_pede_fab_days_da", "PedestalMcg" },
        { "fm_pede_mcg_days_da", "Pedestal" },
        { "fm_diaph_manu_days_da", "WeldedDiaphragm" },
        { "fm_base_frame_man_days_da", "BasePlate" },
        { "fm_grid_val_dia_mcg_days_da", "GridValveDiaphragmMcg" },
        { "fm_gear_box_fab_days_da", "GearboxFabrication" },
        { "fm_gear_box_mcg_days_da", "GearboxMcg" },
        { "fm_nozzle_hou_days_da", "NozzleChest" },
        { "fm_nc_housing_mcg_days_da", "NCHousingICPIVMcg" },
        { "fm_pass_out_days_da", "PassOutManifoldMcg" },
        { "fm_gbc_mcg_days_da", "GBC" },
        { "fm_esv_column_mcg_days_da", "ESVColumnMcg" },
        { "fm_esv_body_mcg_days_da", "ESVBodyMcg" },
        { "fm_laby_holder_days_da", "LabyHolderMcg" },
        { "fm_tv_column_days_da", "TVColumnMcg" },
        { "fm_panting_plate_days_da", "PantingPlateManufacturing" },
        { "fm_packing_mcg_days_da", "PackingMcg" },
        { "fm_tv_internals_days_da", "TVInternalsMcg" },
        { "fm_relay_mcg_days_da", "RelayCylinderMcg" },
        { "fm_blade_shield_days_da", "BladeShieldManufacturing" },
        { "fn_sev_internals_days_da", "SEVInternalsMcg" },
        { "fm_c_class_days_da", "CClassRaviComponentsMcg" },
        { "fm_tip_seal_days_da", "TipSealManufacturing" },
        { "fm_remarks_sdt250", "Remarks" },
        { "fm_other_c_class_days_da", "OtherCClassItemsLeadTime" },
        { "fm_long_lead_ors_da", "LongLeadORS" },
        { "k__creator_id", "CreatedBy" },
        { "creator_id", "CreatedName" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__uuu_record_last_update_user", "UpdatedBy" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" }
    };

    // Hardcoded column mapping for BankMaster table
    private static readonly Dictionary<string, string> BankMasterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "BankMasterID" },
        { "record_no", "RecordNo" },
        { "bm_bank_sdt50", "Bank" },
        { "bm_bank_name_sdt50", "NameOfTheBank" },
        { "bm_ac_type_sdt50", "AcType" },
        { "bm_branch_sdt50", "Branch" },
        { "bm_address_sdt250", "Address" },
        { "bm_acc_no_sdt50", "BankAcNo" },
        { "bm_ifsc_code_sdt50", "IFSCCode" },
        { "bm_swift_code_sdt50", "SwiftCode" },
        { "bm_bsr_code_sdt50", "BSRCode" },
        { "status", "Status" },
        { "k__creator_id", "CreatedBy" },
        { "creator_id", "CreatedName" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__uuu_record_last_update", "UpdatedBy" },
        { "uuu_record_last_update_us", "UpdatedName" },
        { "uuu_record_last_update_da", "UpdatedAt" }
    };


    private static readonly Dictionary<string, string> LetterOfCorrespondenceMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "record_no", "RecordNo" },
            { "ucpt_kindly_attn_sdt120", "KindlyAttention" },
            { "status", "Status" },
            { "uuu_record_last_update_date", "UpdatedAt" },
            { "creator_id", "PrimaveraCreatedId" },
            { "project_id", "ProjectId" },
            { "uuu_user_email", "EmployeeEmail" },
            { "uuu_creation_date", "CreatedAt" },
            { "usmusassigneddeptpd", "EmployeeAssignedDepartment" },
            { "loc_subject_stb500", "Subject" },
            { "loc_ccaddress_txt4000", "CCAddress" },
            { "loc_seqno_sas", "LOCSequenceNumber" },
            { "loc_emailbody_txt4000", "EmailBody" },
            { "loc_employee_name_smn", "EmployeeName" },
            { "uuu_user_title", "EmployeeJobTitle" },
            { "uloc_our_reference_sdt255", "OurReference" },
            { "loc_creator_name_smn", "CreatedName" },
            { "loc_creator_title_smn", "CreatorTitle" },
            { "uuu_record_last_update_user", "PrimaveraUpdatedId" },
            { "id", "LetterOfCorrespondenceId" },
            { "employee_bpk", "EmployeeId" },
            { "loc_reference_sdt120", "Reference" },
            { "ordertransmittalid", "OrderTransmittalId" },
        {"uot_sold_party_dp","CustomerMasterID" },
        {"uot_ship_to_partydp","EndUserID" }
    };

    private static readonly Dictionary<string, string> ContractClearanceMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "record_no", "RecordNo" },
            { "status", "Status" },
            { "creator_id", "CreatedName" },
            { "k__creator_id", "PrimaveraCreatedId" },
        //{"uuu_record_last_update_date","UpdatedAt" },
        {"uuu_record_last_update_user" ,"UpdatedName"},
        {"k__uuu_record_last_update_user","PrimaveraUpdatedId" },
            { "uot_turbine_pd", "TypeofTurbine" },
            { "uot_contract_clnce_sdt120", "ContractClearanceFormat" },
            { "type_of_warranty_pd", "TypeOfWarranty" },
            { "uot_cs_pd", "ServiceType" },
            { "uot_qap_pd", "QAP" },
            { "order_acceptance", "OrderAcceptance" },
            { "turbine_material_code", "TurbineMaterialCode" },
            { "uuu_record_last_update_date", "UpdatedAt" },
            { "k__cc_sel_rec_dpk", "CCRecordSelectionId" },
            { "uuu_creation_date", "CreatedAt" },
            { "uot_ratiing_ia", "TurbineRatingKW" },
            { "ot_date", "OTDate" },
            { "uot_nonstandard_pd", "FrameStandard" },
            { "ucc_schedul_dop", "ScheduledDispatchDate" },
            { "specify_if_non_standard_tb", "FrameNonStandard" },
            { "ucc_kom1", "ProposedKickOffDate" },
            { "comissioning_spares_pd", "TypeofSpares" },
            { "cc_bm_turbine_dop1", "TurbineBillingMonth" },
            { "ugeninstructionmtl4000", "SpecialInstructions" },
            { "cc_bv_turbine_da", "TurbineBillingValue" },
            { "cc_bv_dbo_da", "DBOBillingValue" },
            { "cc_bm_dbo_dop", "DBOBillingMonth" },
            { "k__uot_ship_to_partydp", "EndUserID" },
            { "k__uot_sold_party_dp", "CustomerMasterID" },
            { "project_id", "ProjectId" },
            { "id", "ContractClearanceId" },
        {"ordertransmittalid","OrderTransmittalId" }
    };

    // Hardcoded column mapping for AdditionalOrderBooking table
    private static readonly Dictionary<string, string> AdditionalOrderBookingMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "status", "Status" },
        { "ao_item_des_sdt500", "ItemDescription" },
        { "ao_order_status_pd", "OrderStatus" },
        { "ao_month_book_dop", "BookingMonth" },
        { "ao_month_bill_dop", "BillingMonth" },
        { "ao_basic_cost_dec", "BasicCostINR" },
        { "ao_total_price_dec", "TotalPriceINR" },
        { "ao_est_profit_margin_dec", "EstimatedProfitMarginPercent" },
        { "ao_saving_value_dec", "SavingValueINR" },
        { "record_no", "RecordNo" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "id", "AdditionalOrderBookingId" },
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" }
    };

    // Hardcoded column mapping for ContractOnHold table
    private static readonly Dictionary<string, string> ContractOnHoldMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "project_id", "ProjectId" },
        { "cr_specify_if_other_sdt250", "SpecifyIfOther" },
        { "uuu_creation_date", "CreatedAt" },
        { "proposed_date_of_revival_pd", "IsProposedDateOfRevival" },
        { "proposed_date_of_revival_dp", "ProposedDateofRevival" },
        { "cr_contract_status_pd", "ContractStatus" },
        { "cr_remarks", "Remarks" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "cr_has_project_put_on_hold", "ProjectOnHoldInPrimevera" },
        { "cr_has_project_locked_sap", "ProjectLockedinSAP" },
        { "cr_reason_for_hold_cancel", "ReasonForHold" },
        { "id", "ContractOnHoldId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "k__creator_id", "PrimaveraCreatedId" },
        {"creator_id" ,"CreatedName"},
        { "cr_project_leader_remarks", "ProjectLeaderRemarks" },
        { "cr_project_hod_remarks", "ProjectHODRemarks" },
        { "cr_contract_status_pd2", "UpdatedContractStatus" },
        { "cr_reason_for_cancellation_", "ReasonForCancellation" },
        { "proposed_date_of_revival_d1", "DateOfRevival" },
        { "cr_remarks1", "ProjectRevivalRemarks" },
        { "proposed_date_of_revival_d2", "DateOfCancellation" },
        { "cr_rev_date_dop", "RevisedDeliveryDate" },
        { "cr_project_hod_remarks1", "ProjectHODReview" }
    };

    // Hardcoded column mapping for LCReview table
    private static readonly Dictionary<string, string> LCReviewMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "lcr_rev_no_da", "CurrentRevision" },
        {"record_no","RecordNo" },
        {"status","Status" },
        { "lcr_log_remarks_sdt2000", "LogisticsRemarks" },
        { "lcr_pm_rmks", "PMRemarks" },
        { "lcr_intor_rmks", "InitiatorRemarks" },
        { "lcr_draft_final_pd", "DraftFinal" },
        { "lc_type_pd", "LCType" },
        { "lcr_comm_doc_nam_pd", "CommercialDocumentName" },
        { "lcr_dat_of_comp", "ExpectedDateOfShipmentCompletion" },
        { "lcr_val_da", "Value" },
        { "lcr_curncy_pd", "Currency" },
        { "uot_incoterms_pd", "INCOTerms" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "project_id", "ProjectId" },
        { "id", "LCReviewId" }
    };

    // Hardcoded column mapping for LCReview_NotesObservation table
    private static readonly Dictionary<string, string> LCReviewNotesObservationMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "LCReviewNotesObservationId" },
        { "record_id", "LCReviewId" },
        { "notes", "Notes" },
        { "lcr_obsv_2000", "Observations" },
        { "lcr_resp_dept_pd", "ResponsibleDepartment" },
        { "lcr_stat_pd", "ObservationsStatus" }
    };

    // Hardcoded column mapping for InitialCashPlan table
    private static readonly Dictionary<string, string> InitialCashPlanMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "id", "InitialCashPlanId" },
        { "record_no", "RecordNo" },
        { "ci_tax_duties_da", "TaxesDutiesPercent" },
        { "ci_tax_amt_da", "TotalTaxableAmountINR" },
        { "ci_net_payable_da", "SubTotalINR" },
        { "ci_tcs_amt_da", "TCSAmount" },
        { "ci_tot_net_payable_da", "TotalNetPayableAmountINR" },
        { "ci_tcs_da", "TCSPercent" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "cf_icp_for_rb", "InitialCashFlowPlanFor" }
    };

    private static readonly Dictionary<string, string> MinutesOfMeetingMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MOMID" },
        { "record_no", "RecordNo" },
        { "project_id", "ProjectID" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "status", "status" },
        { "process_status", "ProcessStatus" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "ugenmeetlctnnametxt120", "MeetingLocationName" },
        { "ugenmeetdatestrttimedp", "MeetingDateTime" },
        { "ugenmeetminsupk", "MeetingCoordinatorName" },
        { "k__ugenmeetminsupk", "MeetingCoordinatorPrimaveraId" },
        { "ugenmeettypemmpd", "MeetingType" },
        { "ugensbjctssdt32", "Subject" }
    };

    // Hardcoded column mapping for MOM_Attendees child table
    private static readonly Dictionary<string, string> MOMAttendeesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_id", "MOMID" },
        { "ummattndmn", "AttendeeName" },
        { "ugeninattdmapd", "IsPresent" }
    };

    // Hardcoded column mapping for MOM_Minutes child table
    private static readonly Dictionary<string, string> MOMMinutesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_id", "MOMID" },
        { "ugendetailmtl4000", "Details" },
        { "mm_ass_to_sdt50", "AssignedTo" },
        { "ugenduedatedo", "DueDate" },
        { "ugennamessdt120", "Title" }
    };

    // Hardcoded column mapping for PaymentSupply table
    private static readonly Dictionary<string, string> PaymentSupplyMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "uec_payment_terms", "PaymentTerms" },
        { "ci_mil_per_da", "MilestonePercent" },
        { "ci_type_of_payment_ipd", "TypeOfPayment" },
        { "ib_supply_da", "SupplyValue" },
        { "ot_exchange_rate_da", "ExchangeRate" },
        { "uot_cur1_pd", "Currency" },
        { "total_amount_da_ot", "TotalAmountINR" },
        { "record_id", "OrderTransmittalID" },
        { "id", "PaymentSupplyID" },
        { "project_id", "Projectid" }
    };

    // Hardcoded column mapping for LiquidatedDamage table
    private static readonly Dictionary<string, string> LiquidatedDamageMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "ot_ldclse_sdt500", "LDClause" },
        { "ot_type_ld_pd", "LiquidatedDamageType" },
        { "ot_ld_max_pd", "IsAmountPercent_Maximum" },
        { "ot_ld_min_pd", "IsAmountPercent_Minimum" },
        { "ot_min_amout_dec", "MinimumAmount_INRofPercent" },
        { "ot_max_amount_dec", "MaximumAmount_INRofPercent" },
        { "ot_min_1_dec", "MinimumPercent" },
        { "ot_max_5_dec", "MaximumPercent" },
        { "ot_ld_min_da", "MinimumAmountINR" },
        { "ot_ld_max_da", "MaximumAmountINR" },
        { "uuu_li_last_update_date", "UpdatedAt" },
        { "others", "LiquidatedDamageOthersSpecify" },
        { "record_id", "OrderTransmittalID" },
        { "id", "LiquidatedDamageID" },
        { "Projectid", "ProjectID" }
    };

    // Hardcoded column mapping for PaymentENC table
    private static readonly Dictionary<string, string> PaymentENCMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "ot_pay_per_da", "PaymentInPercent" },
        { "ot_milestone_pd", "TypeOfPayment" },
        { "ot_payment_sdt500", "PaymentTerms" },
        { "total_amount_da_ot", "TotalAmountINR" },
        { "record_id", "OrderTransmittalID" },
        { "id", "PaymentENCID" },
        { "Projectid", "ProjectID" }
    };

    // Hardcoded column mapping for OrderTransmittal_Notes table
    private static readonly Dictionary<string, string> OrderTransmittalNotesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        //{ "NotesId", "Id" },
        { "record_id", "OrderTransmittalID" },
        {"id","NotesId" },
        { "project_id", "ProjectId" },
        { "notes", "Notes" }
    };

    // Hardcoded column mapping for OrderTransmittal_ContractualDeliveryMapping table
    private static readonly Dictionary<string, string> OrderTransmittalContractualDeliveryMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ot_remarks4_sdt250", "ContractualCommencementRemarks" },
        { "ot_contrac_comm_dop", "ContractualCommencement" },
        { "contractual_delivery_date_d", "CDContractualDeliveryDate" },
        { "ot_remarks5_sdt250", "ContractualDeliveryDateRemarks" },
        { "uot_con_comm_dop", "CDContractualCommissioningDate" },
        { "ot_remarks6_sdt250", "ContractualCommissioningRemarks" },
        { "ot_contr_del_weeks_da", "ContractualDeliveryWeeks" },
        { "ot_contr_del_ia", "ContractualDeliveryDifferenceDays" },
        { "ot_contr_comm_weeks_da", "ContractualCommissioningWeeks" },
        { "ot_contr_comm_ia", "ContractualCommissioningDifferenceDays" },
        { "duration_in_months", "DurationInMonths" },
        { "remarks_tb14", "DurationInMonthsRemarks" },
        { "ot_approval_req3_msi", "ContractDeliveryApprovalRequiredFor" },
        { "terms_of_delivery_incoterms", "TermsOfDeliveryIncoterms" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_DefectLiability table
    private static readonly Dictionary<string, string> OrderTransmittalDefectLiabilityMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ot_original_warr_rb", "OriginalWarranty" },
        { "ot_remarks_sdt250", "OriginalWarrantyRemarks" },
        { "ot_repair_warr_rb", "RepairReplacementWarranty" },
        { "ot_remarks2_sdt250", "RepairReplacementWarrantyRemarks" },
        { "ot_defect_ext_warr_rb", "LatentDefectWarranty" },
        { "ot_remarks3_sdt250", "LatentDefectWarrantyRemarks" },
        { "remarks_tb3", "DefectLiabilityRemarks" },
        { "ot_approval_req2_msi", "DefectLiabilityApprovalRequiredFor" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_FormatsReviewed table
    private static readonly Dictionary<string, string> OrderTransmittalFormatsReviewedMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ot_formats_rev_msi", "FormatsReviewed" },
        { "Others2", "FormatsReviewedOthers" },
        { "remarks_tb12", "FormatsReviewedRemarks" },
        { "uot_commercial_pk", "CommercialManager" },
        { "ucc_review_dop", "CommercialReviewDate" },
        { "ugenremarkstb4000", "CommercialReviewSummaryRemarks" },
        { "uot_exec_review_ldt", "ProposalExecutiveReview" },
        { "uot_prosal_ldt", "ProposalHeadReview" },
        { "uot_prosal_1ldt", "ProposalHeadApprovalReview" },
        { "uot_markting_head_rev", "MarketingHeadReview" },
        { "uot_proect_heads_ldt", "ProjectsHeadReview" },
        { "uot_project_pan_pd", "HasProjectNumberCreatedInSAP" },
        { "remarks_tb13", "PMRemarks" },
        { "uot_project_pan_pd1", "HasProjectPlanCreated" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_GoverningLaw table
    private static readonly Dictionary<string, string> OrderTransmittalGoverningLawMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ot_governing_law_rb", "GoverningLaw" },
        { "ot_others_specify2_sdt250", "GoverningLawOthersSpecify" },
        { "ot_dispute_resol_rb", "DisputeResolution" },
        { "ot_others_specify3_sdt250", "DisputeResolutionOthersSpecify" },
        { "ot_no_arbitrators_rb", "NoOfArbitrators" },
        { "ot_others_specify4_sdt250", "NoOfArbitratorsOthers" },
        { "ot_language_rb", "Language" },
        { "ot_others_specify5_sdt250", "LanguageOthers" },
        { "jurisdiction_pd", "Jurisdiction" },
        { "if_not_bangalore_please_sp", "IfNotBanglorePleaseSpecify" },
        { "remarks_tb10", "GoverningLawRemarks" },
        { "ot_approval_req4_msi", "GoverningLawApprovalRequiredFor" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_ScopeOfWork table
    private static readonly Dictionary<string, string> OrderTransmittalScopeOfWorkMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "supply_rb", "Supply" },
        { "ot_transportation_rb", "Transportation" },
        { "transit_insurance_rb", "SOWTransitInsurance" },
        { "custom_clearance_rb", "CustomClearance" },
        { "crane_at_site_rb", "CraneAtSite" },
        { "ear_insurance_rb", "EARInsurance" },
        { "storage_at_site_rb", "StorageAtSite" },
        { "eandc_complete_rb", "EandCComplete" },
        { "eandc_supervision_rb", "EandCSupervision" },
        { "eandc_lodging_and_boardrb", "EandCLodgingandBoarding" },
        { "pgtr_rb", "PGTR" },
        { "uot_statu_app_pd", "StatutoryApprovals" },
        { "remarks_tb16", "ScopeOfWorkRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_RelevantTerms table
    private static readonly Dictionary<string, string> OrderTransmittalRelevantTermsMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "storage_clause_rb", "StorageClause" },
        { "ot_remarks7_sdt250", "StorageClauseRemarks" },
        { "confidentiality_rb", "Confidentiality" },
        { "ot_remarks9_sdt250", "ConfidentialityRemarks" },
        { "consequential_damage_inderb", "ConsequentialDamage" },
        { "ot_remarks5_sdt250", "ConsequentialDamageRemarks" },
        { "intellectual_patent_inderb", "Intellectual" },
        { "ot_remarks10_sdt250", "IntellectualRemarks" },
        { "risk_and_title_transfer_rb", "RiskandTitleTransfer" },
        { "ot_remarks11_sdt250", "RiskandTitleTransferRemarks" },
        { "force_majeure_rb", "ForceMajeure" },
        { "ot_remarks12_sdt250", "ForceMajeureRemarks" },
        { "insurance_rb", "Insurance" },
        { "ot_remarks13_sdt250", "InsuranceRemarks" },
        { "suspension_cancellatio_rb", "SuspensionAndCancellation" },
        { "ot_remarks14_sdt250", "SuspensionAndCancellationRemarks" },
        { "termination_clause_rb", "TerminationClause" },
        { "ot_remarks15_sdt250", "TerminationClauseRemarks" },
        { "ot_purchaser_oblig_rb", "PurchaserObligation" },
        { "ot_remarks16_sdt250", "PurchaserObligationRemarks" },
        { "any_other_critical_clasuerb", "AnyOtherCriticalClause" },
        { "ucc_relevent_others_sdt120", "AnyOtherCriticalClauseOthers" },
        { "ot_approval_req_msi", "RelevantTermsApprovalRequiredFor" },
        { "remarks_tb11", "RelevantTermsRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_PerformanceGuarantee table
    private static readonly Dictionary<string, string> OrderTransmittalPerformanceGuaranteeMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "standard_frame_rb", "StandardFrame" },
        { "development_frame_rb", "DevelopmentFrame" },
        { "availability_confirmationrb", "RDPerformanceConfirmation" },
        { "specific_approvals_rb", "SpecificApproval" },
        { "remarks_tb8", "PGDocumentsRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_LiquidatedDamages table
    private static readonly Dictionary<string, string> OrderTransmittalLiquidatedDamagesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ucc_liquid_dam_cb01", "LDLateDelivery_fivepercentCV" },
        { "ld_15_ed_approval_rb", "LDLateDelivery_EDApproval" },
        { "ld1015_hod_approval_rb", "LDLateDelivery_HODApproval" },
        { "remarks_tb4", "LDLateDelivery_Remarks" },
        { "ucc_liquid_dam_CB02", "LDLateDelivery_TenpercentOfCV" },
        { "ucc_liquid_perf_cb01", "LDPerformance_SteamConsumption" },
        { "ucc_aux_cb2", "LDPerformance_AuxiliaryPower" },
        { "ld_15_ed_approval_rb1", "LDPerformance_EDApproval" },
        { "ld1015_hod_approval_rb1", "LDPerformance_HODApproval" },
        { "remarks_tb5", "LDPerformance_Remarks" },
        { "ucc_liquid_sdt2000", "LDDocumentation_DocumenationSpecify" },
        { "ld_on_documentation_pd", "LDDocumentation_LDOnDocumentation" },
        { "remarks_tb6", "LDDocumentation_Remarks" },
        { "not_exceed_10_of_cv_rb", "MaximumCumulativeLD_TenpercentOfCV" },
        { "ld1015_hod_approval_rb2", "MaximumCumulativeLD_HODApproval" },
        { "uot_if_others11_sdt250", "MaximumCumulativeLD_Others" },
        { "others_specify_tb", "MaximumCumulativeLD_OthersSpecify" },
        { "remarks_tb7", "MaximumCumulativeLD_Remarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_PaymentDetails table
    private static readonly Dictionary<string, string> OrderTransmittalPaymentDetailsMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ucc_payment_terms_rb", "PaymentTerms" },
        { "taxes_and_duties_tb", "TaxesAndDuties" },
        { "notes_tb", "PaymentNotes" },
        { "remarks_tb2", "PaymentRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_TermsOfDelivery table
    private static readonly Dictionary<string, string> OrderTransmittalTermsOfDeliveryMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "CV_10_RB", "LimitsOfLiabilityTenpercentCV" },
        { "equal_to_cv_rb", "LimitsOfLiabilityEqualToCV" },
        { "others_rb2", "LimitsOfLiabilityOthers" },
        { "ucc_ld_others_sdt120", "LimitsOfLiabilityOthersSpecify" },
        { "remarks_tb9", "LimitsOfLiabilityRemarks" },
        { "ot_approval10_rb", "LimitsOfLiabilityApproval" },
        { "payment_mechanism", "PaymentMechanism" },
        { "remarks_tb15", "PaymentMechanismRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for OrderTransmittal_CostSheetReview table
    private static readonly Dictionary<string, string> OrderTransmittalCostSheetReviewMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalId" },
        { "ucc_margin_rb", "MarginAsPerBusinessPlan" },
        { "man_days_charges_rb", "ManDaysCharges" },
        { "bg_cost_included_rb", "BGCostIncludedInTheCostSheet" },
        { "ot_exchange_rate_da", "CSRExchangeRate" },
        { "contingency_3_rb", "ThreePercentContingency" },
        { "ot_ec_rb", "ErectionAndCommissioning" },
        { "agent_commission_pd", "AgencyCommission" },
        { "ot_site_insurance_cost_rb", "SiteInsuranceCost" },
        { "special_expenses_rb", "SpecialExpenses" },
        { "additional_warranty_cost_rb", "AdditionalWarrantyCost" },
        { "ld_cost_short_delivery_rb", "LDCostShortDelivery" },
        { "export_packing_forwardingrb", "ExportPackingandForwardingCost" },
        { "additional_contingency_3_rb", "ThreePercentAdditionalContingency" },
        { "transport_cost_as_per_incrb", "TransportCostAsPerINCOTERMS" },
        { "lc_confirmation_charges_rb", "LCConfirmationCharges" },
        { "insurance_cost_rb", "InsuranceCost" },
        { "lc_discounting_charges_rb", "LCDiscountingCharges" },
        { "remarks_tb", "CostSheetRemarks" },
        { "k__creator_id", "PimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for Risks table
    private static readonly Dictionary<string, string> RisksMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "RiskId" },
        { "k__risk_sel_risk_dpk", "OldRiskId" },
        { "record_no", "RecordNo" },
        { "project_id", "ProjectId" },
        { "risk_supply_status_pd", "SupplyStatus" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "risk_financial_year_pd", "FinancialYear" },
        { "risk_quarter_pd", "Quarter" },
        { "risk_ot_pd", "RiskRatingDuringOTStage" },
        { "risk_type_pd", "Type" },
        { "risk_phase_pd", "Phase" },
        { "risk_category_pd", "Category" },
        { "k__risk_risk_identifier_upk", "RiskIdentifierId" },
        { "risk_risk_identifier_upk", "RiskIdentifierName" },
        { "usmusassigneddeptpd01", "RiskIdentifierDepartment" },
        { "k__risk_risk_owner_upk", "RiskOpportunityOwnerId" },
        { "risk_risk_owner_upk", "RiskOpportunityOwnerName" },
        { "usmusassigneddeptpd02", "RiskOpportunityOwnerDepartment" },
        { "risk_brief_desc_sdt250", "RiskActivityDescription" },
        { "riskbp_risk_description", "RiskDescription" },
        { "risk_causes", "Causes" },
        { "risk_conseq", "Consequences" },
        { "risk_cost_impact_probabilit", "CostImpact" },
        { "probable_case_amount", "ExpectedAmount" },
        { "risk_schedule_impact_pd_yn", "ScheduleImpact" },
        { "risk_schedule_impact_pd", "ScheduleImpactWeeks" },
        { "risk_scope_impact_ynpd", "ScopeImpact" },
        { "risk_quality_impact_ynpd", "QualityImpact" },
        { "urischedchangereqdpd", "ScheduleChangeRequired" },
        { "risk_inherent_likelihood", "InherentLikelihood" },
        { "risk_final_inherent_score1", "FinalInherentScore" },
        { "risk_inherent_impact_pd", "InherentImpact" },
        { "inherent_risk_rating_qualit", "InherentRiskRatingQualitative" },
        { "risk_ttl_risk_rating_pd", "TTLInherentRiskRating" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "status", "Status" }
    };

    // Hardcoded column mapping for Risk_MitigationPlan table
    private static readonly Dictionary<string, string> RiskMitigationPlanMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MitigationPlanId" },
        { "record_id", "RiskId" },
        { "risk_prop_miti_plan_mtb4000", "ProposedMitigationPlan" },
        { "risk_prop_date_miti_dop", "ProposedDateOfMitigation" },
        { "risk_curr_miti_stat_mtb4000", "CurrentStatusOfMitigationPlan" },
        { "risk_revi_date_miti_dop", "RevisedDateOfMitigation" },
        { "uuu_li_last_update_date", "UpdatedAt" }
    };

    // Hardcoded column mapping for PositiveRecall table
    private static readonly Dictionary<string, string> PositiveRecallMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "PositiveRecallId" },
        { "project_id", "ProjectId" },
        { "customermasterid", "CustomerMasterId" },
        { "ordertransmittalid", "OrderTransmittalID" },
        { "k__vendor_select_dp", "VendorId" },
        { "status", "Status" },
        { "record_no", "RecordNo" },
        { "pr_item_code", "ItemCode" },
        { "pr_drawing_number", "DrawingNumber" },
        { "pr_description", "ReasonForPositiveRecallRequest" },
        { "pr_proposed_date_of_closure", "ProposedDateOfClosure" },
        { "pr_quantity", "Quantity" },
        { "pr_closure_action", "ClosureAction" },
        { "uot_tur_framepd", "TurbineFrameType" },
        { "pr_item_description", "ItemDescription" },
        { "pr_dep_pd", "InitiatorDepartment" },
        { "pr_hod_name", "HODName" },
        { "pr_hod_approval_date_new", "ApprovalDate" },
        { "pr_hod_comments", "HODComments" },
        { "k__pr_qa_responsible", "PrimaveraQAResponsibleId" },
        { "pr_qa_responsible", "QAResponsibleName" },
        { "pr_qa_res_coomments", "QAResponsibleComments" },
        { "pr_qa_approval_date", "QAApprovalDate" },
        { "k__pr_hod_name", "PrimaveraHODId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimaverAaUpdatedId" },
        { "k__mp_sel_proj_spk", "SelectProjectId" },
        { "vendor_select_dp", "SelectSupplier" },
        { "turbine_material_code", "TurbineMaterialCode" },
        { "enduserid", "EndUserId" },
        { "pr_closure_action_completed", "IsPRClosureActionCompleted" },
        { "pr_date_of_closure2", "DateOfClosure" },
        { "pr_initiator_comments", "InitiatorComments" },
        { "pr_closure_action_verified", "IsPRClosureActionVerified" },
        { "pr_qa_decision_date", "QADecisionDate" },
        { "pr_qa_comments", "QAComments" }
    };

    // Hardcoded column mapping for MonthlyCashInflow table
    private static readonly Dictionary<string, string> MonthlyCashInflowMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MonthlyCashInflowId" },
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "project_id", "ProjectId" },
        { "ci_planned_date5_dp", "MarketingPlannedDate" },
        { "mci_mar_proj_da", "MarketingProjectedPayment" },
        { "k__mci_sel_rec_pk", "OldMonthlyCashInflowId" },
        { "mci_mar_act_da", "MarketingActualPayment" },
        { "ci_planned_date_dop", "CustomerCarePartsPlannedDate" },
        { "mci_cc_proj_parts_da", "CustomerCarePartsProjectedPayment" },
        { "ci_planned_date2_dop", "CustomerCareServicesPlannedDate" },
        { "mci_cc_proj_serv_da", "CustomerCareServicesProjectedPayment" },
        { "ci_planned_date3_ic_dop", "ICPlannedDate" },
        { "mci_ec_proj_da", "ICProjectedPayment" },
        { "ci_planned_date3_dop", "CustomerCareOMPlannedDate" },
        { "mci_cc_proj_om_da", "CustomerCareOMProjectedPayment" },
        { "ci_planned_date4_dop", "CustomerCareRefurbishingPlannedDate" },
        { "mci_cc_proj_refurb_da", "CustomerCareRefurbishingProjectedPayment" },
        { "ci_planned_date5_dop", "CustomerCareSubsidariesPlannedDate" },
        { "mci_cc_proj_subs_da", "CustomerCareSubsidariesProjectedPayment" },
        { "ci_planned_date6_dop", "CustomerCareAgainstBillingPlannedDate" },
        { "mci_cc_proj_bill_da", "CustomerCareAgainstBillingProjectedPayment" },
        { "mci_cc_proj_da", "CustomerCareProjectedPayment" },
        { "mci_cc_act_da", "CustomerCareActualPayment" },
        { "mci_bal_coll_da", "MarketingBalanceToBeCollected" },
        { "mci_sum_bal_coll_da", "CustomerCareBalanceToBeCollected" },
        { "mci_ec_bal_da", "ICBalanceToBeCollected" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatorId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for MonthlyCashInflow_CustomerCare table
    private static readonly Dictionary<string, string> MonthlyCashInflowCustomerCareMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MCI_CustomerCareId" },
        { "record_id", "MonthlyCashInflowId" },
        { "mci_pay_agnst_pd", "PaymentFor" },
        { "ci_actual_date1_dop", "ActualDate" },
        { "mci_actual_coll_da", "ActualPayment" },
        { "mci_proj_da", "ProjectedPayment" }
    };

    // Hardcoded column mapping for MonthlyCashInflow_MarketingAdvance table
    private static readonly Dictionary<string, string> MonthlyCashInflowMarketingAdvanceMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MCI_MarketingAdvanceId" },
        { "record_id", "MonthlyCashInflowId" },
        { "mp_prj_segmnt", "SegmentOfProject" },
        { "ci_1exchange_da", "ExchangeRate" },
        { "ci_1eq_amt_da", "EquivalentAmountINR" },
        { "mac_date_credit_dop", "DateOfCredit" },
        { "ci_actual_date_dop", "ActualDate" },
        { "mac_mode_of_payment_sdt50", "ModeOfPayment" },
        { "ci_bank_name_sdt250", "BankName" },
        { "ci_remarks_mtb4000", "Remarks" },
        { "ordertransmittalid", "OrderTransmittalID" },
        { "customermasterid", "CustomerMasterId" },
        { "k__ci_sel_proj_spk", "ProjectId" },
        { "mci_actual_coll_da", "ActualPayment" },
        { "mci_proj_da", "ProjectedPayment" },
        { "ci_status1_pd", "Status" }
    };

    // Hardcoded column mapping for MonthlyCashInflow_ICDue table
    private static readonly Dictionary<string, string> MonthlyCashInflowICDueMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MCI_ICDueId" },
        { "record_id", "MonthlyCashInflowId" },
        { "ci_project_name_sdt50", "ProjectName" },
        { "ci_actual_date_dop", "ActualDate" },
        { "mci_actual_coll_da", "ActualPayment" },
        { "ci_status1_pd", "Status" },
        { "mci_proj_da", "ProjectedPayment" }
    };

    // Hardcoded column mapping for MonthlyCashInflow_ICNotDue table
    private static readonly Dictionary<string, string> MonthlyCashInflowICNotDueMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MCI_ICNotDueId" },
        { "record_id", "MonthlyCashInflowId" },
        { "ci_project_name_sdt50", "ProjectName" },
        { "ci_actual_date_dop", "ActualDate" },
        { "mci_actual_coll_da", "ActualPayment" },
        { "ci_status1_pd", "Status" },
        { "mci_proj_da", "ProjectedPayment" }
    };

    // Hardcoded column mapping for HandoverProtocol table
    private static readonly Dictionary<string, string> HandoverProtocolMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "HandoverProtocolId" },
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "k__uot_sold_party_dp", "CustomerMasterID" },
        { "k__uot_ship_to_partydp", "EndUserID" },
        { "uot_nonstandard_pd", "FrameStandard" },
        { "uot_order_type_pd", "OrderType" },
        { "uot_tur_framepd", "TurbineFrameType" },
        { "specify_if_non_standard_tb", "FrameNonStandard" },
        { "turbine_serial_no", "TurbineSerialNo" },
        { "turbine_material_code", "TurbineMaterialCode" },
        { "commissioned_status_pd", "CommissioningStatus" },
        { "commissioning_date", "CommissioningDate" },
        { "warranty_completion_date", "WarrantyCompletionDate" },
        { "hp_sap_bills_status_pd", "IsSAPbillingStatusAttached" },
        { "hp_full_payment_pd", "IsFullPaymentCollected" },
        { "hp_no_rem_sdt250", "ActionPlanCollectionIfNo" },
        { "hp_list_issues_ynpd", "ListOfOpenIssue" },
        { "hp_list_commit_ynpd", "ListOfOpenCommitments" },
        { "hp_details_spares_ynpd", "DetailsOfSparesSupply" },
        { "hp_spec_ser_off_sdt250", "SpecialServicesOffered" },
        { "hp_value_spec_sdt250", "ValueForSpecialServices" },
        { "hop_bg_status_pd", "BGStatus" },
        { "hop_bg_exp_dop", "BGExpiry" },
        { "hop_bg_clm_prd_dop", "BGClaimPeriod" },
        { "hp_proj_lock_sap_cb", "IsProjectLockedinSAP" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "hp_value_cost_da", "Value_cost" },
        { "sc_sc_a_status_pd", "SiteClosureAStatus" },
        { "sc_planned_delivery_dp", "PlannedDateOfCompletion" },
        { "es_remarks10_sdt250", "SiteClosureARemarks" },
        { "sc_sc_b_stat_pd", "SiteClosureBStatus" },
        { "sc_pl_date_sc_b_dop", "PlannedDateOfCommissioning" },
        { "es_remarks11_sdt250", "SiteClosureBRemarks" },
        { "sc_sc_c_status_pd", "SiteClosureCStatus" },
        { "sc_pl_date_sc_c_dop", "PlannedDateOfPGtest" },
        { "es_remarks12_sdt250", "SiteClosureCRemarks" },
        { "sc_sc_d_status_pd", "SiteClosureDStatus" },
        { "sc_planned_date_of_expiry", "PlannedDatEOfWarrantyExpiry" },
        { "es_remarks13_sdt250", "SiteClosureDRemarks" },
        { "status", "Status" },
        { "uot_tur_rating_da", "TurbineRatingKW" },
        { "hp_cont_person1_sdt50", "CustomerContactPerson" },
        { "uot_designation3_sdt50", "CustomerDesignation" },
        { "uue_user_contactphone", "CustomerContactPhone" },
        { "hp_cont_person2_sdt50", "EndUserContactPerson" },
        { "ec_des_sdt50", "EndUserDesignation" },
        { "ec_cont_per_smn", "EndUserContactPhone" },
        { "record_no", "RecordNo" },
        { "uot_drive_pd", "DrivenEquipmentName" },
        { "contractclearanceid", "ContractClearanceId" },
        { "format_no", "FormatNumber" }
    };

    // Hardcoded column mapping for HandoverProtocolLineItem table
    private static readonly Dictionary<string, string> HandoverProtocolLineItemMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "HandoverProtocolLineItemId" },
        { "record_id", "HandoverProtocolId" },
        { "uuu_name", "Name" },
        { "hop_type_of_doc_pd", "TypeOfDocument" },
        { "group_type", "Isgroupby" },
        { "uuu_li_last_update_date", "UpdatedAt" },
        { "uuu_file_repo_path", "FilePath" }
    };

    // Hardcoded column mapping for InternalAudit table
    private static readonly Dictionary<string, string> InternalAuditMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "InternalAuditId" },
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "project_id", "ProjectId" },
        { "ia_audit_number", "AuditNumber" },
        { "ia_audit_type", "AuditType" },
        { "ia_audit_name", "AuditName" },
        { "ia_reference", "Reference" },
        { "ia_function_pd", "InternalFunction" },
        { "ia_section_pd", "Section" },
        { "supp_city", "City" },
        { "supp_state_new", "State" },
        { "supp_country_pd", "Country" },
        { "ia_scheduled_start_date", "ScheduledStartDate" },
        { "ia_scheduled_finish_date", "ScheduledFinishDate" },
        { "k__vendor_select_dp", "VendorId" },
        { "dev_supp_name_sdt250", "SupplierName" },
        { "ia_auditee_upk", "AuditeeName" },
        { "ia_ext_aud_sdt250", "ExternalAuditee" },
        { "ia_auditor_upk", "AuditorName" },
        { "ia_coauditor_upk", "CoAuditorName" },
        { "ia_planned_finish_date", "PlannedFinishDate" },
        { "ia_audit_comment", "AuditComment" },
        { "k__ia_auditee_upk", "PrimaveraAuditeeId" },
        { "k__ia_auditor_upk", "PrimaveraAuditorId" },
        { "k__ia_coauditor_upk", "PrimaveraCoAuditorId" },
        { "vendor_select_dp", "SelectSupplier" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "ia_auditee_type_pd", "AuditeeType" },
        { "ia_supp_type_pd", "SupplierType" },
        { "ia_planned_start_date", "PlannedStartDate" },
        { "ia_actual_start_date", "ActualStartDate" },
        { "ia_actual_finish_date", "ActualFinishDate" }
    };

    // Hardcoded column mapping for InternalAudit_Notes table
    private static readonly Dictionary<string, string> InternalAuditNotesMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "InternalAuditNotesId" },
        { "record_id", "InternalAuditId" },
        { "notes", "Notes" },
        { "uuu_li_last_update_date", "UpdatedAt" }
    };

    // Hardcoded column mapping for InternalAudit_ObservationDetails table
    private static readonly Dictionary<string, string> InternalAuditObservationDetailsMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "InternalAudit_ObservationDetailsId" },
        { "record_id", "InternalAuditId" },
        { "ia_observation_number_da", "ObservationNumber" },
        { "ia_observation_df", "Observation" },
        { "ia_observation_type_pd", "ObservationType" },
        { "ia_reference_clause", "ReferenceClause" },
        { "ia_rcca_required_ipd", "IsRCCARequired" },
        { "ia_rcca_required_date_df", "RCCARequiredDate" },
        { "ia_correction_done_df", "CorrectionDone" },
        { "ia_actual_corr_com_date_df", "ActualCorrectiveCompletionDate" },
        { "ia_root_cause_analysis_df", "RootCauseAnalysis" },
        { "ia_corrective_action_tak_df", "CorrectiveActionTaken" },
        { "ia_actual_cor_act_com_da_df", "ActualCorrectionActionCompletionDate" },
        { "ia_verification_dec_ipd", "VerificationDecision" },
        { "ia_veri_deci_date_df", "VerificationDecisionDate" },
        { "ia_reason_for_ver_dec_df", "ReasonForVerificationDecision" },
        { "ia_reference_document", "ReferenceDocument" },
        { "ia_requirement_df", "Requirement" },
        { "ia_evidence_df", "Evidence" },
        { "uuu_li_last_update_date", "UpdatedAt" }
    };

    // Hardcoded column mapping for OTRisk table
    private static readonly Dictionary<string, string> OTRiskMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OTRiskId" },
        { "project_id", "ProjectId" },
        { "ordertransmittalid", "OrderTransmittalId" },
        { "status", "Status" },
        { "record_no", "RecordNo" },
        { "otr_cost_rating_pd", "CostOverrunRiskRating" },
        { "otr_cost_impact_pd", "CostOverrunImpact" },
        { "otr_con_del_rating_pd", "ContractualDeliveryRiskRating" },
        { "otr_con_del_impact_pd", "ContractualDeliveryImpact" },
        { "otr_payment_rating_pd", "CommercialTermsRiskRating" },
        { "otr_payment_impact_pd", "CommercialTermsImpact" },
        { "otr_crs_rating_pd", "CustomerRelationshipRiskRating" },
        { "otr_crs_impact_pd", "CustomerRelationshipImpact" },
        { "otr_financial_rating_pd", "FinancialHealthRiskRating" },
        { "otr_financial_impact_pd", "FinancialHealthImpact" },
        { "otr_tg_rating_pd", "AgreedPerformanceRiskRating" },
        { "otr_tg_impact_pd", "AgreedPerformanceImpact" },
        { "otr_comm_terms_rating_pd", "WarrantyTermsRiskRating" },
        { "otr_comm_terms_impact_pd", "WarrantyTermsImpact" },
        { "otr_m_cost_rating_pd", "CostOverrunRiskRatingPM" },
        { "otr_m_cost_impact_pd", "CostOverrunImpactPM" },
        { "otr_m_con_del_rating_pd", "ContractualDeliveryRiskRatingPM" },
        { "otr_m_con_del_impact_pd", "ContractualDeliveryImpactPM" },
        { "otr_m_payment_rating_pd", "CommercialTermsRiskRatingPM" },
        { "otr_m_payment_impact_pd", "CommercialTermsImpactPM" },
        { "otr_m_crs_rating_pd", "CustomerRelationshipRiskRatingPM" },
        { "otr_m_crs_impact_pd", "CustomerRelationshipImpactPM" },
        { "otr_m_financial_rating_pd", "FinancialHealthRiskRatingPM" },
        { "otr_m_financial_impact_pd", "FinancialHealthImpactPM" },
        { "otr_m_tg_rating_pd", "AgreedPerformanceRiskRatingPM" },
        { "otr_m_tg_impact_pd", "AgreedPerformanceImpactPM" },
        { "otr_m_comm_terms_rating_pd", "WarrantyTermsRiskRatingPM" },
        { "otr_m_comm_terms_impact_pd", "WarrantyTermsImpactPM" },
        { "k__creator_id", "PrimeveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimeveraUpdatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };

    // Hardcoded column mapping for UserList table
    private static readonly Dictionary<string, string> UserListMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "UserId", "UserId" },
        { "UserName", "UserName" },
        { "FirstName", "FirstName" },
        { "LastName", "LastName" },
        { "Email", "Email" },
        { "FullName", "FullName" }
    };

    // Hardcoded column mapping for MonthlyActualCollection table
    private static readonly Dictionary<string, string> MonthlyActualCollectionMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "project_id", "ProjectId" },
        { "id", "MonthlyActualCollectionId" },
        { "status", "Status" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_creation_date", "CreatedAt" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" }
    };


    // Hardcoded column mapping for InitialCashFlowPlan table (child of InitialCashPlan)
    private static readonly Dictionary<string, string> InitialCashFlowPlanMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "InitialCashFlowPlanId" },
        { "record_id", "InitialCashPlanId" },
        { "uuu_li_last_update_date", "UpdatedAt" },
        { "ci_planned_monthh_dp", "PlannedMonth" },
        { "ci_pay_status_pd", "PaymentStatus" },
        { "ci_tax_amt_da", "TotalTaxableAmount" },
        { "ci_full_tcs_tbc_ynipd", "IsFullTaxToBeCollected" },
        { "ci_tcs_applicable_ynpd", "IsTCSApplicable" },
        { "ci_milestone_tax_tbc_ynipd", "IsMilestoneTaxToBeCollected" },
        { "amount", "Amount" },
        { "ci_tcs_amt_da", "TCSAmount" },
        { "ci_net_pay_li_da", "SubTotalINR" },
        { "ci_tot_net_payable_da", "TotalNetPayableAmountINR" },
        { "otid", "OrderTransmittalId" },
        { "projectid", "ProjectId" }
    };

    // Hardcoded column mapping for MonthlyPlanning table
    private static readonly Dictionary<string, string> MonthlyPlanningMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "amount", "TotalEquivalentAmountINR" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "id", "MonthlyPlanningId" },
        { "project_id", "ProjectId" }
    };

    // Hardcoded column mapping for MonthlyPlanningLineItem table (child of MonthlyPlanning)
    private static readonly Dictionary<string, string> MonthlyPlanningLineItemMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "mp_prj_segmnt", "ProjectSegment" },
        { "uot_sales_order_sdt120", "SupplySalesOrderNo" },
        { "ci_pay_agnt1_pd", "PaymentAgainst" },
        { "ci_planned_date1_dp", "PlannedDate" },
        { "ci_baseline_planned_amt_da", "PlannedAmount" },
        { "ci_1exchange_da", "ExchangeRate" },
        { "ci_1eq_amt_da", "EquivalentAmountINR" },
        { "ci_status1_pd", "Status" },
        { "ci_lc_type1_pd", "LCType" },
        { "ci_risk_rating1_pd", "RiskRating" },
        { "mci_remarks1_mdt4000", "Remarks" },
        { "uuu_li_last_update_date", "UpdatedAt" },
        { "id", "MonthlyPlanningLineItemId" },
        { "k__ci_sel_proj_spk", "ProjectId" },
        { "record_id", "MonthlyPlanningId" },
        { "k__ci_sel_ini_cf_plan_dpk", "InitialCashPlanId" },
        { "Otid", "OrderTransmittalId" },
        { "k__ci_sel_month_lidp", "InitialCashFlowPlanId" },
        { "ci_sel_month_lidp", "PlannedMonth" }
    };





    // Hardcoded column mapping for OrderTransmittal tables (applies to tables with "OrderTransmittal" prefix)
    private static readonly Dictionary<string, string> OrderTransmittalMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "OrderTransmittalID" },
            { "record_no", "RecordNo" },
            { "process_status", "ProcessStatus" },
            { "status", "Status" },
            { "project_id", "ProjectID" },
        {"k__uppr_project_pk","ProjectManagerId" },

            { "k__customer_contacts_dp", "CustomerContactID" },
            { "k__customer_contacts_dp1", "CustomerContactID2" },
            { "k__customer_contacts_dp3", "EndUserContactID" },
            { "k__customer_contacts_dp4", "EndUserContactID2" },

            { "uot_bus_sec_pd", "BusinessSector" },
            { "uot_if_others11_sdt250", "OthersBusinessSector" },

            { "uot_order_type_pd", "OrderType" },
            { "ot_order_class_pd", "EPCorDirect" },

            { "ci_order_date_dop", "OrderDate" },
            { "uot_po_num_sdt120", "PurchaseOrderNumber" },
            { "uot_po_date_dop", "PODate" },
            { "uot_contract_sdt120", "ContractNumber" },
            { "uot_contract_date_dop", "ContractDate" },
            { "uot_agreement_sdt120", "AgreementNumber" },
            { "uot_agreement_date_dop", "AgreementDate" },

            { "uot_first_adv_pd", "FirstAdvanceReceived" },
            { "uot_miles_dop", "ReceiptOfFirstAdvancePaymentDate" },
            { "uot_frquency_pd", "Frequency" },

            { "uot_con_deliv_dop", "ContractualDeliveryDate" },
            { "contractual_commissioning", "ContractualCommissioningDate" },

            { "ot_currency_smn", "Currency" },
            { "ot_exchange_rate_da", "ExchangeRate" },

            { "uot_cs_pd", "ServiceType" },
            { "ib_supply2_da", "SupplyValue" },
            { "turnkey_value_ca", "Turnkey" },
            { "supervision_ca", "Supervision" },
            { "free_manday_supervision_da", "FreeMandaySupervision" },
            { "charges_free_manday_da", "ChargesAfterFreeMandays" },

            { "scope_of_spares", "ScopeOfSpares" },
            { "ot_total_order_value2_da", "OrderValue" },
            { "spares_value_ca", "SpareValue" },
            { "ci_order_valueinr_da", "OrderValueINR" },
            { "ot_order_supp_da", "OrderValueSupply" },
            { "ot_order_ec_da", "OrderValueEandC" },
            { "ot_order_supp_inr_da", "OrderValueSupplyINR" },
            { "ot_order_ec_inr_da", "OrderValueEandCINR" },

            { "uot_costsheet_ynpd", "CostSheetAttached" },
            { "agent_commission_pd", "AgentCommission" },
            { "uot_incoterms_pd", "INCOTerms" },
            { "uot_gst_pd", "GST" },

            { "uot_sea_worthy_packing_pd", "ScopeOfSeaworthyPacking" },
            { "uot_marine_ins_pd", "MarineInsurance" },
            { "uot_taxes_pd", "TaxesDutiesSpecify" },
            { "ci_tax_duties_da", "TaxesAndDutiesPercent" },
            //old
            //{ "k__uot_sold_party_dp", "CustomerMasterID" },
            //{ "k__uot_ship_to_partydp", "EndUserID" },
            //new 
             { "k__uot_sold_party_dp", "CustomerMasterID" },
            { "k__uot_ship_to_partydp", "EndUserID" },

            { "uot_sales_order_sdt120", "SupplySaleOrderno" },
            { "uot_sales_order_ec_sdt120", "ECSaleOrderNo" },
            { "uot_c_number_sdt120", "CProjectNumber" },
            { "k__uot_bpcreator_bc", "OldOTId" },

            { "uot_spl_notes_sdt2000", "SpecialNotes" },
            { "uot_spl_notes1_sdt2000", "SpecialNotesCustomerInformation" },

            { "ot_type_order_pd", "TypeOfOrder" },
            { "uot_site_insurance_pd", "SiteInsurance" },
            { "transit_insurance", "TransitInsurance" },
            { "uot_compre_insurance_pd", "ComprehensiveInsurance" },
            { "uot_fright_pd", "ScopeOfFrieght" },

            { "uot_spcy_sdt_250", "LimitIfAgreed" },
            { "uot_statu_app_pd", "StatutoryApproval" },

            { "otr_cost_rating_pd", "CostOverrunRiskRating" },
            { "otr_cost_impact_pd", "CostOverrunImpact" },
            { "otr_con_del_rating_pd", "ContractualDeliveryRiskRating" },
            { "otr_con_del_impact_pd", "ContractualDeliveryImpact" },
            { "otr_payment_rating_pd", "CommercialTermsRiskRating" },
            { "otr_payment_impact_pd", "CommercialTermsImpact" },

            { "otr_crs_rating_pd", "CustomerRelationshipRiskRating" },
            { "otr_crs_impact_pd", "CustomerRelationshipImpact" },

            { "otr_financial_rating_pd", "FinancialHealthRiskRating" },
            { "otr_financial_impact_pd", "FinancialHealthImpact" },

            { "otr_tg_rating_pd", "AgreedPerformanceRiskRating" },
            { "otr_tg_impact_pd", "AgreedPerformanceImpact" },

            { "otr_comm_terms_rating_pd", "WarrantyTermsRiskRating" },
            { "otr_comm_terms_impact_pd", "WarrantyTermsImpact" },

            { "uot_transmittaltype_pd", "TransmittalTypeID" },
        { "uot_loi_sdt120", "LetterOfIntentNumber" },
{ "uot_loi_date_dop", "LOIDate" },
{ "uot_comp_name2_sdt250", "CompanyNameConsultant" },
{ "uot_contact_name1_sdt250", "ContactPersonNameConsultant" },
{ "uot_designation3_sdt50", "DesignationConsultant" },
{ "uot_email3_tb120", "EmailConsultant" },
{ "uot_phone3_sdt50", "PhoneConsultant" },
{ "ugenfaxtxt16", "FaxConsultant" },
{ "uot_cons_citytxt50", "CityConsultant" },
{ "uot_india_consu_states_pd", "StateProvinceConsultant" },
{ "uot_consu_countrypd", "CountryConsultant" },
{ "uot_state2", "OtherStateProvinceConsultant" },
{ "uot_types_serv_pd", "TypesOfServicesEandC" },
{ "ot_mob_pd", "MobileCraneFacilityEandC" },
{ "uot_erection_crane_pd", "EotCraneFacilityEandC" },
{ "uot_erection_pd", "ErectionCraneEandC" },
{ "uot_conev_pd", "ConveyanceForEngineerEandC" },
{ "uot_unloading_pd", "UnloadingAtSiteEandC" },
{ "uot_grouting_pd", "GroutingEandC" },
{ "uot_grout_pd", "GroutingMaterialSupplyEandC" },
{ "uot_storage_pd", "StorageAtSiteEandC" },
{ "uot_const_pd", "ConstructionPowerWaterEandC" },
{ "uot_erection_cable_pd", "ErectionCableAndBaseEandC" },
{ "comissioning_spares_pd", "TypeOfSparesEandC" },
{ "uot_spares_desc_sdt500", "DescriptionEandC" },
{ "uot_additiona_sdt250", "AdditionalScopeConditions" },

{ "type_of_warranty_pd", "TypeOfWarranty" },
{ "others_please_specify1", "OtherTypeOfWarranty" },
{ "replaced_parts_warranty_pd", "ReplacedPartsWarranty" },
{ "uot_amb_temp_da", "AmbientTemperature" },
{ "uot_temp9_unit_pd", "AmbientTemperatureUnitID" },
{ "uot_temp_da", "TemperatureRiseDeltaT" },
{ "uot_temp12_unit_pd", "TemperatureRiseDeltaTUnitID" },
{ "ot_so_min_int", "TemperatureMin" },
{ "ot_so_max_int", "TemperatureMax" },

{ "uot_humidity_da", "RelativeHumidityPercent" },
{ "uot_altitude_da", "AltitudeAboveMSLMetres" },

{ "uot_earth_zone_pd", "EarthquakeZone" },
{ "uot_if_others8_sdt250", "EarthquakeZoneOther" },

{ "uot_ed_da", "ElectricalDesign" },
{ "uot_temp10_unit_pd", "ElectricalDesignUnitID" },

{ "uot_iap_dp", "InstrumentAirPressure" },
{ "uot_unit25_pd", "InstrumentAirPressureUnitID" },

{ "uot_cooling_water_pd", "CoolingWater" },
{ "uot_supply_da", "CWSupplyTemperature" },
{ "uot_temp11_unit_pd", "CWSupplyTemperatureUnitID" },
{ "uot_supply_pressure_da", "CWSupplyPressure" },
{ "uot_unit26_pd", "CWSupplyPressureUnitID" },

{ "uot_design_pres_sdt120", "DesignPressure" },
{ "uot_unit28_pd", "DesignPressureUnitID" },

{ "uot_presuure_drop_sdt120", "PressureDrop" },
{ "uot_unit27_pd", "PressureDropUnitID" },

{ "uot_motor_eff_pd", "MotorEfficiency" },
{ "uot_if_others6_sdt250", "MotorEfficiencyOther" },

{ "uot_main_pd", "GeneratedVoltageRating" },
{ "uot_if_others7_sdt250", "GeneratedVoltageOther" },

{ "uot_variation_da", "VariationCWPercent" },
{ "uot_variation2_da", "VariationFreqPercent" },
{ "uot_com_da", "CombinedVariationPercent" },

{ "uot_aux_power_pd", "AuxiliaryVoltageRating" },
{ "ele_others_23", "AuxiliaryVoltageOther" },

{ "uot_environment_pd", "Environment" },
{ "uot_if_others5_sdt250", "EnvironmentOther" },

{ "uot_scopepd", "ScopeForCivil" },
{ "uuu_creation_date", "CreatedAt" },
{ "creator_id", "CreatedName" },
 {"k__creator_id","PrimaveraCreatedId" },

{ "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
{ "uuu_record_last_update_date", "UpdatedAt" },
{ "uuu_record_last_update_user", "UpdatedName" },
    };

    //Hardcoded column mapping for CustomerMaster
    private static readonly Dictionary<string, string> CustomerMasterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "id", "CustomerID" },
            { "record_no", "RecordNo" },

            { "uuu_record_last_update_date", "UpdatedAt" },
            { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
            {"uuu_record_last_update_user","UpdatedName" },

            { "status", "Status" },

            { "k__creator_id", "PrimaveraCreatedId" },
             { "creator_id", "CreatedName" },
            { "uuu_creation_date", "CreatedAt" },

            { "ucm_comp_name_sdt120", "CompanyName" },
            { "uot_sold_party_code_sdt120", "CompanyCode" },

            { "uuu_proj_phone", "Phone" },
            { "phone_number004", "WorkPhone" },

            { "uveemailtb120", "Email" },
            { "uuu_user_fax", "FaxNumber" },

            { "uvetaxidtb16", "GST" },
            { "uvelicensenotb16", "LicenseNo" },

            { "uot_shipping_pd", "StateIndia" },
            { "uot_state", "StateOther" },

            { "ugencitytxt50", "City" },
            { "ugencountrypd", "Country" },

            { "ucm_url_hp", "CompanyURL" },

            // Composite address – concatenate in code
            {
                "ugenaddress1txt120+ugenaddress2txt120+ugenaddress3txt120",
                "Address"
            }
        };

    //Hardcoded column mapping for CustomerContacts
    private static readonly Dictionary<string, string> CustomerContactMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
           { "id", "CustomerContactID" },
            { "record_no", "RecordNo" },
            { "title", "Title" },
 
            { "uuu_record_last_update_date", "UpdatedAt" },
            { "k__creator_id", "PrimaveraCreatedId" },
            { "uuu_creation_date", "CreatedAt" },
        { "creator_id","CreatedName"},

            { "uircntctfstnmtb", "ContactName" },
            { "customermasterid", "CustomerID" },

            { "uuu_proj_city", "City" },
            { "uuu_user_state", "State" },
            { "ugenzipcodetxt16", "ZipPostalCode" },
            { "ugencountrypd", "Country" },

            { "uot_designation3_sdt50", "Designation" },

            { "uue_user_contactphone", "ContactPhone" },
            { "uuu_user_workphone", "WorkPhone" },

            { "uveemailtb120", "Email" },

            // Composite address – must be concatenated in code
            {
                "uaddress1txt120+uaddress2txt120+ugenaddress3txt120",
                "Address"
            }
            // You can expand this dictionary with actual Excel column names
        };

    //Hardcoded column mapping for VendorMaster
    private static readonly Dictionary<string, string> VendorMasterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
             { "id", "VendorID" },
            { "record_no", "RecordNo" },

            { "uuu_record_last_update_date", "UpdatedAt" },
            { "uuu_record_last_update_user", "PrimaveraUpdatedIdName" },
            { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },

            { "process_status", "ProcessStatus" },
            { "status", "StatusID" },

            { "creator_id", "PrimaveraCreatedIdName" },
            { "k__creator_id", "PrimaveraCreatedId" },

            { "vendor_master_vendor", "VendorName" },
            { "vendor_master_con_person", "ContactPerson" },
            { "vendor_master_manu_add", "ManufacturingAddress" },
            { "vendor_master_con_number", "ContactNumber" },

            { "uuu_creation_date", "CreatedAt" },
            { "vendor_master_code", "VendorCode" }
        };

    private static readonly Dictionary<string, string> MechanicalDBOMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "id", "MechanicalDBOId" },
            { "record_no", "RecordNo" },
            { "process_status", "ProcessStatus" },
            { "status", "Status" },
            { "project_id", "ProjectId" },
        {"k__ot_sel_ot_rec_bpp","OrderTransmittalID" },
            { "uot_others_411", "TubeSheetsAdditionalDetails" },
            { "uot_others_410", "ShellAdditionalDetails" },
            { "uot_start_ejector_pd", "StartupEjector" },
            { "uot_others_412", "TubesAdditionalDetails" },
            { "uot_water_pd", "WaterBoxes" },
            { "uot_others_4181", "QuantityAdditionalDetails" },
            { "uot_ttlscope7_pd", "AdditionalBOPScope" },
            { "uot_shell_interpd", "ShellOfInterAfterCondenser" },
            { "uot_unit29_pd", "ExhaustPressureUnitID" },
            { "uot_tubes_pd", "Tubes" },
            { "uot_spl_notes3_sdt2000", "SpecialNotes" },
            { "uot_plugging_pd", "PluggingMargin" },
            { "uot_others_408", "WaterBoxesAdditionalDetails" },
            { "uot_others_407", "StartupEjectorAdditionalDetails" },
            { "uot_ejector_pd", "EjectorNozzle" },
            { "uot_mech_outlet_temp_pd", "CWOutletTemperature" },
            { "uot_others_409", "BafflesAdditionalDetails" },
            { "uot_design_pressure_pd", "CWDesignPressure" },
            { "uot_conden_scope_pd", "CondenserScope" },
            { "uot_condensing_sdt120", "CondensingCapacity" },
            { "gland_scope", "GlandVentCondenserScope" },

            { "uuu_creation_date", "CreatedAt" },

            { "ele_others_11", "TubesOfInterAfterCondenserAdditionalDetails" },
            { "uot_add_bop_pd", "AdditionalBOP" },
            { "ele_others_08", "MainEjectorAdditionalDetails" },
            { "uot_fouling_pd", "FoulingFactor" },
            { "uot_during_start_pd", "EjectionSystemDuringStartup" },
            { "uot_vel_pd", "CWVelocity" },
            { "condensate_shell", "GlandVentShell" },
            { "glans_tubes", "GlandVentTubes" },
            { "uot_flow_rating_pd", "FlowRating" },
            { "remarks_tb1", "CondensorRemarks" },
            { "pressure10_pd", "PressureUnitID" },

            { "uot_others_429", "AdditionalBOPAdditionalDetails" },
            { "uot_others_426", "CleanlinessFactorAdditionalDetails" },
            { "uot_auxiliary_steam_da", "AuxiliarySteamTemperature" },
            { "uot_others_425", "FoulingFactorAdditionalDetails" },
            { "uot_tube_sheets_pd", "TubeSheets" },
            { "uot_others_422", "CWOutletTemperatureAdditionalDetails" },
            { "uot_others_421", "CWSupplyPressureAdditionalDetails" },
            { "uot_auto_gland_pd", "AutoGlandSealingSystem" },
            { "uot_others_424", "PluggingMarginAdditionalDetails" },
            { "exhaust_pressure_condensyst", "ExhaustPressure" },
            { "uot_others_423", "CWInletTemperatureAdditionalDetails" },
            { "uot_baffles_pd", "Baffles" },
            { "uot_others_420", "CWDesignPressureAdditionalDetails" },
            { "uot_condensate_pd", "CondensateExtractionPumpScope" },
            { "uot_auxilary_steam_da", "AuxiliarySteamPressure" },
            { "uot_rated_diff_head_pd", "RatedDifferentialHead" },
            { "uot_others_419", "CWVelocityAdditionalDetails" },
            { "uot_others_418", "HotelWellRetentionTimeAdditionalDetails" },
            { "lp_gland_sealing", "LPGlandSealingAndDesuperheater" },
            { "uot_others_415", "FlowRatingAdditionalDetails" },
            { "uot_others_417", "RatedDifferentialHeadAdditionalDetails" },
            { "uot_others_416", "MaterialOfCasingAdditionalDetails" },

            { "uot_ttlscope3_pd", "EjectionSystemScope" },
            { "uot_tubesheet_pd", "TubesSheetOfInterAfterCondenser" },
            { "ot_select_project_sp", "CloneProjectId" },

            { "remarks_tb13", "AuxiliaryRemarks" },
            { "remarks_tb15", "EjectionRemarks" },
            { "remarks_tb11", "CondensateRemarks" },
            { "remarks_tb12", "GlantVentRemarks" },

            { "ambient_temperature1", "AmbientTemperature" },
            { "uot_safety_condensor_pd", "SafetyDeviceForCondenser" },
            { "condensate_sheet", "GlandVentTubesSheet" },
            { "uot_mech_cleanli_pd", "CleanlinessFactor" },
            { "ele_others_50", "GlandVentShellAdditionalDetails" },
            { "uot_materail_pd", "MaterialOfCasing" },
            { "ele_others_51", "GlandVentTubesAdditionalDetails" },
            { "ele_others_52", "GlandVentTubesSheetAdditionalDetails" },
            { "uot_roto_meter_pd", "Rotometer" },
            { "msparameter_scope", "MSParameterGlandSealingEjectionSystemScope" },
            { "uot_inetr_pd", "InterAfterCondenser" },
            { "uot_relief_valve_pd", "ReliefValve" },
            { "uot_qty_pd", "Quantity" },
            { "uot_gland_sealing_ms", "GlandSealing" },
            { "temperature10_pd", "TemperatureUnitID" },
            { "uot_tubesof_inter_pd", "TubesOfInterAfterCondenser" },
            { "uot_temp8_unit_pd", "AmbientTemperatureUnitID" },
            { "ele_others_20", "EjectionSystemDuringStartupAdditionalDetails" },
            { "ele_others_21", "EjectionSystemForContinuousAdditionalDetails" },
            { "gland_blower", "Blower" },
            { "uot_cw_inlet_pd", "CWInletTemperature" },
            { "uot_main_ejector_pd", "MainEjector" },
            { "ot_vac_bre_pd", "VacuumBreakerValve" },
            { "uot_shell_pd", "Shell" },
            { "uot_condesning_type_pd", "Type" },
            { "uot_hotel_pd", "HotelWellRetentionTime" },
            { "uot_mec_dump_pd", "DumpProvision" },
            { "uot_supply_pressure_pd", "CWSupplyPressure" },
            { "uot_cross_overduct_pd", "CrossOverduct" },
            { "uot_for_continuous_pd", "EjectionSystemForContinuous" }
        };

    // Hardcoded column mapping for BPAttachments table
    private static readonly Dictionary<string, string> BPAttachmentMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "BPAttachmentID" },
        { "parent_type", "RecordNo" },
        { "project_id", "ProjectID" },
        { "file_name", "FileName" },   // file_name maps to FileName; FilePath is handled via node_path below
       // { "parent_id", "OrderTransmittalRecordID" },  // Conditionally mapped based on parent_type (only when parent_type = 'uxot2')
        { "create_date", "CreatedAt" },
        { "create_by", "PrimaveraCreatedId" },
        {"parent_id","UnifierAttchmentID" },
        {"version","version" },
        { "node_path","FilePath"},   // node_path maps to FilePath (prevents duplicate when file_name also present)
        {"doc_id","doc_id" }
    };
    // Hardcoded column mapping for BPAttachments when AttachmentRecordType = "Comment"
    private static readonly Dictionary<string, string> BPAttachmentCommentMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        //{ "id", "BPAttachmentID" },
        { "parent_type", "RecordNo" },
        { "project_id", "ProjectID" },
        { "file_name", "FileName" },   // Note: file_name maps to both FileName and FilePath (handled in MatchColumnsForBPAttachments)
        //{ "parent_id", "BPCommentRecordID" },  // For Comment type, parent_id maps to BPCommentRecordID
        { "upload_date", "CreatedAt" },
        { "upload_by", "PrimaveraCreatedId" },
         {"parent_id","UnifierAttchmentID" }

    };

    // Hardcoded column mapping for BPAttachments when AttachmentRecordType = "OrderTransmittal"
    private static readonly Dictionary<string, string> BPAttachmentOTMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        //{ "id", "BPAttachmentID" },
        { "parent_type", "RecordNo" },
        { "project_id", "ProjectID" },
        { "file_name", "FileName" },   // Note: file_name maps to both FileName and FilePath (handled in MatchColumnsForBPAttachments)
        { "parent_id", "OrderTransmittalRecordID" },  // For OrderTransmittal type, parent_id maps to OrderTransmittalRecordID
        { "upload_date", "CreatedAt" },
        { "upload_by", "PrimaveraCreatedId" }
    };

    // Hardcoded column mapping for BPComments table
    private static readonly Dictionary<string, string> BPCommentsMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "detail_id", "BPCommentsID" },
        { "project_id", "ProjectID" },
        { "file_name", "CompanyID" },
        { "content", "Comments" },
        { "creatorid", "PrimaveraCreatedId" },
        { "upload_by", "UserName" },
        //{ "parent_object_id", "" },
        { "lastmodified", "UpdatedAt" },
        {"parent_object_id","UnifierCommentsID" },
        {"parent_object_type","Attachments" }
    };
    private static readonly Dictionary<string, string> BPCommentsOrderTransmittalRecordMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "BPCommentsID" },
        { "project_id", "ProjectID" },
        { "file_name", "CompanyID" },
        { "content", "Comments" },
        { "creatorid", "PrimaveraCreatedId" },
        { "upload_by", "UserName" },
        { "parent_object_id", "OrderTransmittalRecordID" },
        { "lastmodified", "UpdatedAt" },
            {"header_id","UnifierCommentsID" }
    };

    // Hardcoded column mapping for Turbine table
    private static readonly Dictionary<string, string> TurbineMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "TurbineID" },
            { "record_no", "RecordNo" },
            { "title", "Title" },
            { "due_date", "DueDate" },
            { "end_date", "EndDate" },

            { "uuu_record_last_update_date", "UpdatedAt" },
            { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
            { "uuu_record_last_update_user", "UpdatedName" },

            { "process_status", "ProcessStatus" },
            { "status", "StatusId" },

            { "k__creator_id", "PrimaveraCreatedId" },
            { "creator_id", "CreatedName" },

            { "project_id", "ProjectId" },

            { "uot_gearbox_pd", "GearBoxTypeID" },
            { "uot_eff_cons_pd", "EfficiencyId" },

            { "ec2_remarks2_mdt4000", "RemarksCoupling" },

            { "uot_control_oil_filter_pd", "ControlOilFilterId" },
            { "uot_hmbd_pd", "HMBDSubmittedId" },
            { "uot_drive_pd", "DrivenEquipmentId" },
            { "uot_dirct_rot_da", "RotationDirectionID" },
            { "uot_tube_pd", "TubeMOCId" },
            { "uot_mec_over_pd", "OverHeadTankId" },

            { "uuu_dm_publish_path", "uuu_dm_publish_path" },
            { "uuu_creation_date", "CreatedAt" },

            { "uot_min_power_da1", "MinLoadExtraction" },


            { "ec2_remarks9_mdt4000", "Remarks_Gearbox" },
            { "uot_noisepd", "GearBox_NoiseLevelID" },

            { "uot_docu_ms", "DocumentationID" },

            { "ele_others_06", "AnyOtherPointsAdditionalDetails" },
            { "ele_others_07", "TubeSheetsAdditionalDetails" },

            { "ot_others_specify_sdt250", "ManufacturingStandardOthersSpecify" },

            { "uot_mech_sb_sdt250", "StatorBlades" },
            { "gearbox_scope", "GearBoxScope" },

            { "uot_secondary_gear_pd", "SecondaryGearBoxID" },
            { "uot_driven_eqptpd", "SecondaryGBDrivenEqId" },

            { "uot_required_mos_pd", "IfRequiredMOCId" },

            { "uot_others_433", "FoulingFactorAdditionalDetails" },
            { "uot_mech_oil_heaters_pd", "OilHeatersId" },
            { "uot_others_432", "SSTypeAdditionalDetails" },

            { "uot_turbine_pd", "TypeOfTurbineId" },
            { "uot_others_435", "OtherSpecify_Type1" },
            { "uot_others_434", "TubeMOCAdditionalDetails" },

            { "uot_type_inlet_pd", "InletOrientationId" },
            { "uot_others_431", "OilCentrifugeAdditionalDetails" },
            { "uot_others_430", "IfRequiredCapacityAdditionalDetails" },

            { "uot_mech_bs_sdt250", "uot_mech_bs_sdt250" },

            { "ele_others_64", "PluggingMarginAdditonalDetails" },

            { "ot_ss_type_pd", "SSTypeId" },
            { "uot_manu_sta_pd", "ManufacturingStandardID" },

            { "uot_if_others10_sdt250", "NoiseLeveOthersSpecify" },

            { "uot_mech_reduction_pd", "ReductionID" },
            { "ec1_remrks6_mdt4000", "GovernorRemarks" },

            { "uot_if_others9_sdt250", "DrivenEquipmentOthersSpecify" },

            { "k__ot_sel_ot_rec_bpp", "OrderTransmittalID" },

            { "uot_material_pd", "MaterialOfConstruction" },
            { "uot_mech_ffpd", "FoulingFactorId" },

            { "uot_ttlscope6_pd", "GovernorScope" },
            { "uot_mech_casings_sdt250", "Casings" },

            { "uot_nonstandard_pd", "FrameId" },
            //{ "ot_others_specify6_sdt250", "SSTypeAdditionalDetails" },

            { "uot_turbine_details_scope_p", "ScopeId" },
            { "uot_vendor_pd", "VendorListID" },

            { "uot_type_drive_pd", "MOPDriveId" },
            { "uot_mech_accoustic_pd", "AcousticHoodId" },
            { "uot_bearing_pd", "BarringGearID" },

            { "prc1_remarks2_mtb400", "LubeOilRemarks" },

            { "uot_couplg_type_pd", "Type1Id" },
            { "uot_oil_cooler_pd", "OilCoolerId" },

            { "uot_foot_print_replcement_p", "FootPrintReplacementId" },

            { "ec1_remrks5_mdt4000", "Remarks" },

            { "uot_min_load_da", "MinLoadBleed" },

            { "uot_noise_tg_pd", "NoiseLevelID" },



            { "uot_ttlscope8_pd", "HighSpeedScopeId" },

            { "uot_mech_oil_filter_pd", "OilFilterId" },

            { "uot_ratiing_ia", "RatingKW" },

            { "uot_245others_sdt120", "OthersText" },

            { "uot_mech_dirty_pd", "DirtyOilTankId" },

            { "ot_others_specify7_sdt250", "OtherSpecify_Type2" },

            { "uot_mec_vapour_extrr_pd", "VapourExtractorId" },

            { "uot_mech_amot_pd", "AMOTTCVId" },

            { "uot_mech_tubesheet_pd", "TubeSheetsId" },

            { "uot_others_437", "ShortCircuitFactorAdditionalDetails" },
            { "uot_others_436", "GearBox_NoiseLevel_AdditionalDetails" },

            { "low_speed_coupling_type", "Type2Id" },

            { "uot_type_exhst_pd", "ExhaustOrientationId" },

            //{ "uot_service_pd", "ServiceFactorAdditionalDetails" },
            { "uot_others_438", "ServiceFactorAdditionalDetails" },

            { "uot_governorpd", "Governor" },

            { "uot_mangficant_pd", "ShortCircuitFactorID" },

            { "uot_margin_pd", "PluggingMarginId" },

            { "uot_ttlscope5_pd", "LubeOilScopeId" },

            { "uot_drns_ms", "DrawingsID" },

            { "specify_if_non_standard_tb", "NonStandardFrame" },

            { "uot_lube_oil_piping_pd", "LubeOilPipingId" },

            { "uot_mec_points_ms", "AnyOtherPointsId" },

            { "ec2_remarks5_mdt4000", "MaterialRemarks" },

            { "uot_ttl_scope_pd", "PrimarySecondaryGBId" },

            { "uot_spl_notes2_sdt2000", "SpecialNotes" },

            { "uot_capacity_pd", "IfRequiredCapacityId" },

            { "uot_oil_centre_pd", "OilCentrifugeId" },

            { "uot_qap_pd", "QAPID" },

            { "uot_mech_rotor_sdt250", "Rotor" },

            { "low_speed_coupling_scope_pd", "LowSpeedScopeId_pd" },

            { "lubetype", "LubeTypeId" },

            { "uot_mech_rb_sdt250", "RotorBlades" }
    };

    public static readonly Dictionary<string, string> ElectricalInstrumentationDBOMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
     {
         { "id","ElectricalInstrumentationDBOID" },
         { "record_no","RecordNo" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "uuu_record_last_update_date", "UpdatedAt" },

         { "process_status","ProcessStatus" },
         { "status","Status" },
          { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
         { "project_id","ProjectId" },
         { "uot_tvm_mounting_pd","TVMMountingID" },
         { "uot_panel_qty_pd","PanelQtyID" },
         { "ele_others_01","ACBRating_OthersSpecify" },
         { "ele_others_02","Battery_TypeOfCharger_OthersSpecify" },
         { "others_plc","TCP_Others" },
         { "ele_others_03","NumberOfProbes_OthersSpecify" },
         { "mcs_panel_part_of_grp_pd","MCSPannelPartOfGRPID" },
         { "uot_export_cb","IsMeteringCumSynchronizedPanelPartOfACB" },
         { "ip_rating_turbinecontrol","TCP_IPRatingID" },
         { "lt_powercable_moc","LTPowerCableMOCID" },
         { "ucc_export_cost7_cb","Printer" },
         { "uot_elec_rating_da","RatedPower_KW" },
         { "uot_avr_panel_scopepd2alt","AlternatorScopeID" },
         { "uot_quantity15_da","UPSQuantity" },
         { "uot_dc_motor_panle_ms","DCMotorStartPanelID" },
         { "uot_avr_panel_scopepd2meter","SynchronizingPanelScopeID" },
         { "ttl_rated_pf","RatedPFID" },
         { "uot_avr_panel_scopepd2","AVRPanelScopeID" },
         { "ip_rating_dcmotor","DC_Motor_IPRatingID" },
         { "synchronization_grid","SynchronizationGridID" },
         { "ot_length_sdt50","HTPowerCableLength" },
         { "uot_elec_redundancy_ms","TCP_RedundancyID" },
         { "uot_acb_cum_relay_scope_pd2","ACBPanelScopeID" },
         { "uot_hmi_software_for_conf","HMISoftwareForLoadSharingID" },
         { "ip_rating_ng","Transformer_IPRatingID" },
         { "uot_control_mode_pd","ControlModeID" },
         { "uot_avr_panel_scopepd","TurbineGaugePanelScopeID" },
         { "uot_mcc_incomer_qty_pd","IncomerQtyID" },
         { "uot_metering_ct_pt_accuracy","MeteringCTPTAccuracyID" },
         { "uot_alt_enclo_pd","EnclosureID" },
         { "lascpt_panel","LASCPTPanelPartOfBreakerPanelID" },
         { "ttl_if_others_specify","Standard_OthersSpecify" },
         { "uot_bus_bar_material_pd2","SwitchGear_BusBarMaterialID" },
         { "uwarmake1tb120","sync_MakeID" },
         { "uot_others_tb2","RelayRemarks" },
         { "turbine_guage_panel_type","TurbineGaugePanelTypeID" },
         { "uot_spl_notes4_sdt2000","SpecialNotes" },
         { "elec_ot_rmk6_mtb4000","LASCPTRemarks" },
         { "turbovisor_scope","TurbovisorySystemScopeID" },
         { "ttl_standard_pd","StandardID" },
         { "uot_metering_ct_accuracy","MeteringCTAccuracyID" },
         { "ttl_cooler_pd","CoolerConfigurationID" },
         { "uot_resistor_current_limiti","ResistorCurrentLimitingCapacityID" },
         { "ip_rating_avr","Avr_IPRatingID" },
         { "dcmotor_scope","DCMotorStartPanelScopeID" },
         { "ttl_stand_by_exitation","avr_StandByExcitationID" },
         { "uot_acb_cum_relay_quantity","NumberOfBreakersID" },
         { "elec_ot_rmk3_mtb4000","ACBRemarks" },
         { "uot_ct_pt_pd","CTPTID" },
         { "ot_length2_sdt50","BusDuctLength" },
         { "ot_bdt_pd","BusDuctTypeID" },
         { "earth_moc","EarthMOCID" },
         { "ot_elec_cab_moc_pd","InstrumentCablingMOCID" },
         { "uot_type_of_construction_pd","TypeOfConstructionID" },
         { "lt_power","LTPowerCablingID" },
         { "elec_ot_rmk14_mtb4000","CablesRemarks" },
         { "cooler_certification_pd","CoolerCertificationID" },
         { "uot_tvm_accuracy_pd","TVMAccuracyID" },
         { "uot_fault_pd","LASCPT_FaultRatingID" },
         { "ot_remarks24_sdt250","LTPowerCabling_OthersSpecify" },
         { "ot_remarks26_sdt250","HTPowerCabling_OthersSpecify" },
         { "uot_specify_qty_pd","TransducerQuantityID" },
         { "ip_rating_meter","IPRatingID" },
         { "ele_others_405","NumberOfBreakerForSynchOthersSpecify" },
         { "uot_htbraker_panel_pd","SwitchGearTypeID" },
         { "electronic_overspeed","ElectronicOverspeedProtectionID" },
         { "uot_vibration_pd","VibrationMeasurementsTypeID" },
         { "uot_type_pd","Battery_TypeID" },
         { "ot_remarks27_sdt250","BusDuct_OthersSpecify" },
         { "bus_bar_moc","BusBarMOCID" },
         { "comments_tb1","LTPowerCabling_Comments" },
         { "ip_rating_lascpt","LASCPT_IPRatingID" },
         { "uot_steps_pd","DC_Motor_StepsID" },
         { "ttl_tubes_moc","CoolerTubesMOCID" },
         { "ttl_continuous_overload","ContinuousOverloadID" },
         { "elec_ot_rmk2_mtb4000","AVRRemarks" },
         { "ucc_export_cost8_cb","UPS" },
         { "ng_paneltype","TypeOfPanelID" },
         { "uot_avr_panel_scopepd2plc","TurbineControlPanelScopeID" },
         { "type_of_control_panel","TypeOfControlPanelID" },
         { "ttl_mounting","CoolerMountingID" },
         { "ttl_if_others_specify14","ResistorCurrentLimitingCapacity_OthersSpecify" },
         { "uot_impulse_tube_material","Guage_ImpulseTubeMaterialID" },
         { "uot_temp_pd","TemperatureRaiseID" },
         { "uot_duty_rating_pd","DutyRatingID" },
         { "uot_sil_rating_pd","TCP_SILRatingID" },
         { "uot_earthing_pd","EarthingID" },
         { "uot_acdb_pd","ACDBID" },
         { "ttl_slip_ring","SlipRingID" },
         { "ttl_if_others_specify109","MakeOfVMS_OthersSpecify" },
         { "ttl_noise_level","NoiseLevelID" },
         { "elec_ot_rmk5_mtb4000","TransformerRemarks" },
         { "rated_voltage3","RatedVoltage_KV" },
         { "elec_ot_rmk12_mtb4000","DC_Motor_Remarks" },
         { "uot_number_pd","NumberOfBreakerForSynchID" },
         { "elec_ot_rmk8_mtb4000","MotorControlRemarks" },
         { "ht_cable_moc","HTPowerCableMOCID" },
         { "uot_incommer_pd","DC_Motor_IncomerID" },
         { "ip_rating_battery","Battery_IPRatingID" },
         { "ttl_tests","TestsID" },
         { "uot_add_pd","AdditionalProbesID" },
         { "ttl_cooling_method","CoolingMethodID" },
         { "voltage_alternator2","VoltageAlternatorID" },
         { "elec_ot_rmk7_mtb4000","SwitchGearRemarks" },
         { "uot_bcc_capc_pd","Battery_CapacityID" },
         { "uuu_creation_date","CreatedAt" },
         { "elec_ot_rmk4_mtb4000","Metering_Remarks" },
         { "uot_scope_pd2","SwitchGearPanelScopeID" },
         { "elec_ot_rmk13_mtb4000","Turbovisory_Remarks" },
         { "uot_scope_pd1","LASCPTPanelScopeID" },
         { "uot_type_of_synchronizer_pd","TypeOfSynchronizerID" },
         { "uot_number_of_master_module","NumberOfMasterModulesID" },
         { "uot_ct_pd","CTID" },
         { "uot_incomer_type_pd","IncomerTypeID" },
         { "uot_plc_based_instruments","PLCBasedInstrumentsID" },
         { "uot_bbt_pd","BusDuctID" },
         { "cablesbusduct_scope","CablesScopeID" },
         { "uot_instrument_cabling_pd","InstrumentCablingID" },
         { "uot_fault_rating_pd","Transformer_FaultRatingID" },
         { "uot_redundant_ct_pt","GRP_RedundantCTPTInputProtectionID" },
         { "uot_acb_rating_pd","ACBRatingID" },
         { "uot_bcc_volt_pd","Battery_VoltageRatingID" },
         { "number_of_probes","NumberOfProbesID" },
         { "ttl_if_others_specify17","Transformer_FaultRating_OthersSpecify" },
         { "ip_rating_turbineguage","Gauge_IPRatingID" },
         { "uot_scope_pd2batt","BatteryBatteryChargerScopeID" },
         { "uot_make_vms_pd","MakeOfVMSID" },
         { "ht_power","HTPowerCablingID" },
         { "uot_ctpt_pd","CTPTAccuracyID" },
         { "ot_htpct_pd","HTPowerCablingTypeID" },
         { "uot_neutral_isolator","NeutralIsolatorID" },
         { "uot_scope_pd2motor","MotorControlCentreScopeID" },
         { "uot_fault_rating_pd2","SwitchGear_FaultRatingID" },
         { "ucc_export_cost5_cb","Desktop" },
         { "uot_dcdb_pd","DCDBID" },
         { "ttl_insulation_class","InsulationClassID" },
         { "elec_ot_rmk_mtb4000","Remarks" },
         { "alternator_certification_pd","AlternatorCertificationID" },
         { "elec_ot_rmk10_mtb4000","TCP_Remarks" },
         { "uot_relay_ms","AdditionalRelayID" },
         { "uot_relay_type_pd","RelayTypeID" },
         { "ttl_pmg","PMGID" },
         { "uot_metering_pd","MeteringAccuracyID" },
         { "ttl_if_others_specify7","NoiseLevel_OthersSpecify" },
         { "uot_redundant_ct_pt_for","ACB_RedundantCTPTInputProtectionID" },
         { "ot_remarks25_sdt250","ControlCabling_OthersSpecify" },
         { "ttl_design_temperature_pd","DesignTemperatureID" },
         { "uot_avr_panel_scopepd2ng","TransformerPanelScopeID" },
         { "uot_other_electric_sdt2000","OtherElectricalItemsEquipments" },
         { "uot_quantity14_da","PrinterQuantity" },
         { "ttl_temperature_rise","TemperatureRiseID" },
         { "uot_individual_ctpt_pd","IndividualCTPTForTVM" },
         { "ip_rating_motor","MotorControl_IPRatingID" },
         { "uot_redundant_control_pd","RedundantControlID" },
         { "uot_elec_poles_pd","NumberOfPolesID" },
         { "uot_control_cabling_pd","ControlCablingID" },
         { "uot_generator_cb","IsGeneratorRelayPanelPartOfACBPanel" },
         { "make_ms","MakeId" },
         { "ip_rating_switch","SwitchGear_IPRatingID" },
         { "uot_commu_type_pd","TCP_CommunicationTypeID" },
         { "elec_ot_rmk11_mtb4000","GaugeRemarks" },
         { "elec_ot_rmk9_mtb4000","BatteryRemarks" },
         { "ttl_if_others_specify_da","DesignTemperature_OthersSpecify" },
         { "ttl_type","AVRTypeID" },
         { "pqm_pd","PQMID" },
         { "ip_rating_generator","Relay_IPRatingID" },
         { "ip_rating_acb","Acb_IPRatingID" },
         { "ttl_terminal_box_to_suit_pd","TerminalBoxToSuitID" },
         { "uot_discrete_type_pd","TransducerTypeId" },
         { "uot_load_module_pd","LoadSharingModulesScopeID" },
         { "ttl_phase_side_ct_location","PhaseSideCTLocationID" },
         { "ttl_neutra_cr_star_format","NeutralCTStarFormationID" },
         { "ttl_if_others_specify20","LASCPT_FaultRating_OthersSpecify" },
         { "uot_htrating_pd","QuantityRatingID" },
         { "control_cable_moc","ControlCableMOCID" },
         { "uot_bus_bar_material3","MotorControl_BusBarMaterialID" },
         { "ttl_if_others_specify29","Battery_Capacity_OthersSpecify" },
         { "uot_others_9","Earthing_OthersSpecify" },
         { "uot_hmi_laptop_quantity_tb","HMILaptopQuantity" },
         { "uot_number_of_slave_modules","NumberOfSlaveModulesID" },
         { "uot_others_3","TransducerQuantityOthersSpecify" },
         { "uot_mcc_spce_pd","SpecificationID" },
         { "uot_stand_by_exitation","StandByExcitationID" },
         { "uot_others_5","TypeOfConstruction_OthersSpecify" },
         { "uot_quantity13_da","LaptopQuantity" },
         { "uot_others_6","TCP_Specification_OthersSpecify" },
         { "ttl_if_others_specify30","Battery_Type_OthersSpecify" },
         { "uot_softwaregrp_conf_pd","SoftwareGRPConfHMI_ID" },
         { "uot_elect_scope_pd","TCP_SpecificationID" },
         { "uot_level_of_instrumentatio","Guage_LevelOfInstrumentationLogicID" },
         { "ucc_export_cost6_cb","Laptop" },
         { "uot_float_cum_boost_charger","Battery_TypeOfChargerID" },
         { "uot_bus_bar_material_pd","BusBarMaterialID" },
         { "uot_scope_pd1_generator","GeneratorRelayPanelScopeID" },
         { "uot_tvm_type_pd","TVMTypeID" },
        
         { "ele_others_404","TVMMountingOthersSpecify" },
         { "uot_others_10","InstrumentCabling_OthersSpecify" },
         { "uot_bus_bar_material_pd1","Transformer_BusBarMaterialID" },
         { "comments_tb3","HTPowerCabling_Comments" },
         { "ttl_if_others_specify8","CoolerConfiguration_OthersSpecify" },
         { "uot_bus_bar_material1","LASCPT_BusBarMaterialID" },
         { "uot_others_2","Tests_OthersSpecify" },
         { "uot_others_1","InsulationClass_OthersSpecify" },
         { "ttl_if_others_specify28","Battery_VoltageRating_OthersSpecify" },
         { "ttl_if_others_specify26","SwitchGear_FaultRating_OthersSpecify" },
         //{ "uot_others_7","TCP_Remarks" },
         { "ttl_if_others_specify25","MeteringCTPTAccuracy_OthersSpecify" },
         { "ttl_if_others_specify22","SwitchGearType_OthersSpecify" },
         { "ot_remarks22_sdt250","DC_Motor_Steps_OthersSpecify" },
         { "ot_remarks20_sdt250","TurbineGaugePanelType_OthersSpecify" },
         { "ot_remarks23_sdt250","DC_Motor_IPRating_OthersSpecify" },
         { "ttl_if_others_specify1","Enclosure_OthersSpecify" },
         { "ttl_if_others_specify9","CoolerTubesMOC_OthersSpecify" },
         { "ele_others_402","CoolingMethod_OthersSpecify" },
         { "ttl_if_others_specify6","ContinuousOverload_OthersSpecify" },
         { "ucc_export_cost9_cb","Console" },
         { "uot_quantity17_da","ConsoleQuantity" },
         { "uot_quantity12_da","DesktopQuantity" },
         { "uot_others_11","InstrumentCablingMOC_OthersSpecify" },
         { "comments_tb2","ControlCabling_Comments" },
         { "ele_others_65","TemperatureRise_OthersSpecify" },
         { "ttl_if_others_specify15","NeutralIsolator_OthersSpecify" },
         { "ttl_if_others_specify110","TCP_CommunicationType_OthersSpecify" },
         { "ele_others_406","MasterModulesOthersSpecify" },
         { "ttl_if_others_specify10","OthersSpecify" },
         { "ele_others_407","SlaveModulesOthersSpecify" },
         { "ttl_if_others_specify100","RatedPF_OthersSpecify" },
         { "ttl_if_others_specify102","NumberOfBreakers_OthersSpecify" },
         { "ttl_if_others_specify108","Guage_ImpulseTubeMaterial_OthersSpecify" },
         { "ttl_if_others_specify111","TCP_SILRating_OthersSpecify" },
         { "ot_remarks21_sdt250","DC_Motor_Incomer_OthersSpecify" },
         { "ot_sel_ot_rec_bpp","OrderTransmittalID" }
     };

    public static readonly Dictionary<string, string> MonthlyActualCollectionPlannedMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id",                        "MonthlyActCollectionPlannedId" },
        { "k__mac_monthly_planning_dpk", "MonthlyPlanningId" },
        { "k__mac_sel_month_lidp",      "MonthlyPlanningLineItemId" },
        { "ci_actual_amt_da",           "ActualAmount" },
        { "ci_remarks_mtb4000",         "Remark" },
        { "ci_1eq_amt_da",              "EquivalentAmountINR" },
        { "ci_bank_name_sdt250",        "BankName" },
        { "mac_mode_of_payment_sdt50",  "ModeOfPayment" },
        { "mac_actual_date_dop",        "ActualDate" },
        { "mac_date_credit_dop",        "DateOfCredit" },
        { "mac_credited_pd",            "IsCredit" },
        { "uuu_li_last_update_date",    "UpdatedAt" },
        { "otid",                       "OrderTransmittalID" },
        { "record_id",                  "MonthlyActualCollectionId" },
        { "projectid",                  "ProjectId" }
    };


    public static readonly Dictionary<string, string> MonthlyActualUnplannedCollectionMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "ci_lc_type1_pd",              "LcType" },
        { "mp_prj_segmnt",               "SegmentOfProject" },
        { "mac_have_initial_cf_ynpd",    "DoYouHaveInitialCashFlow" },
        { "ci_risk_rating1_pd",          "RiskRating" },
        { "ci_pay_agnt1_pd",             "PaymentAgainst" },
        { "k__ci_sel_ini_cf_plan_dpk",   "InitialCashPlanId" },
        { "id",                          "MonthlyActualUnplannedCollectionId" },
        { "k__ci_sel_proj_spk",          "ProjectId" },
        { "ci_1eq_amt_da",               "EquivalentAmountINR" },
        { "ci_actual_amt_da",            "ActualAmount" },
        { "ci_remarks_mtb4000",          "Remark" },
        { "ci_bank_name_sdt250",         "BankName" },
        { "mac_mode_of_payment_sdt50",   "ModeOfPayment" },
        { "mac_actual_date_dop",         "ActualDate" },
        { "mac_date_credit_dop",         "DateOfCredit" },
        { "mac_credited_pd",             "IsCredited" },
        { "uuu_li_last_update_date",     "UpdatedAt" },
        { "ci_status1_pd",               "Status" },
        { "otid",                        "OrdertransmittalId" },
        { "record_id",                   "MonthlyActualCollectionId" }
    };


    // Hardcoded column mapping for SpecificationRelease table
    public static readonly Dictionary<string, string> SpecificationReleaseMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id",                        "SpecificationReleaseId" },
        { "project_id",                "ProjectId" },
        { "ordertransmittalid",         "OrderTransmittalId" },
        { "status",                    "Status" },
        { "udr_doc_category_pd",       "Category" },
        { "uvr_itemc_pd",              "ItemCategory" },
        { "sr_item_code_sdt50",        "ItemCodeSAP" },
        { "others_coolers",            "ItemCodeSAPOthers" },
        { "quantity",                  "Quantity" },
        { "sr_spares_ynpd",            "Spares" },
        { "sr_if_yes_spec_sdt2000",    "SparesSpecify" },
        { "udoc_resource_name_sdt255", "ResourceName" },
        { "ugenp6plannedstartdop",     "P6PlannedStart" },
        { "ugenp6plannedfinishdop",    "P6PlannedFinish" },
        { "uuu_p6actualstart",         "ActualStart" },
        { "uuu_p6actualfinish",        "ActualFinish" },
        { "sr_spec_comm_sdt2000",      "AnySpecialComment" },
        { "record_no",                 "RecordNo" },
        { "uuu_creation_date",         "CreatedAt" },
        { "k__creator_id",             "PrimaveraCreatedId" },
        { "creator_id",                "CreatedName" },
        { "uuu_record_last_update_date",      "UpdatedAt" },
        { "k__uuu_record_last_update_user",   "PrimaveraUpdatedId" },
        { "uuu_record_last_update_user",      "UpdatedName" },
    };

    // Hardcoded column mapping for SparesOrderTransmittal table
    private static readonly Dictionary<string, string> SparesOrderTransmittalMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "SparesOrderTransmittalID" },
        { "record_no", "RecordNo" },
        { "project_id", "ProjectID" },
        { "usot_ot_type_sdt250", "TransmittalType" },
        { "status", "Status" },
        { "uprot_bpk", "SparesProvisionalOrderReference" },
        { "uexot_de_bp", "XOTRef" },
        { "order_type_spares_pd", "SparesOrderType" },
        { "loi_or_order_for_tb", "LOIorOrderfor" },
        { "turbine_model_and_s_no", "TurbineModelAndSNo" },
        { "spares_quotation_ref", "SAPQuotationReference" },
        { "quotation_reference_date", "QuotationReferenceDate" },
        { "customer_poref", "CustomerPOReference" },
        { "customer_poref_date", "CustomerPOReferenceDate" },
        { "po_receipt_at_ttl", "POReceiptAtTTIL" },
        { "basic_quoted_price", "BasicQuotedPrice" },
        { "final_po_value_accepted", "FinalPOValueAcceptedForSupplyPart" },
        { "discount_accepted_for_suppl", "DiscountAcceptedForSupplyPart" },
        { "discount_accepted_for_job", "DiscountAcceptedForJobWork" },
        { "price_basis", "PriceBasis" },
        { "spares_contractual_delivery", "ContractualDeliveryForSupplyPart" },
        { "contractual_delivery_explai", "ContractualDeliveryExplanation" },
        { "contractual_delivery_for_si", "ScopeOfSiteWork" },
        { "final_price_accepted_for_si", "FinalPriceAcceptedForSiteWork" },
        { "ld_clause_accepted_if_any", "LDClauseAcceptedPercent" },
        { "discount_accepted_for_site_", "DiscountAcceptedForSiteWork" },
        { "transportation_mode", "TransportationMode" },
        { "remarks_regarding_site_work", "RemarksRegardingSiteWork" },
        { "transportation_charges", "TransportationCharges" },
        { "packing_and_forwarding_new", "PackingAndForwardingChargesAcce" },
        { "insurance_charges", "InsuranceCharges" },
        { "taxes_duties", "TaxesAndDuties" },
        { "payment_remarks_tb", "PaymentTermsAcceptedWithCusto" },
        { "rsc_required_delivery_date", "RSCRequiredDeliveryDate" },
        { "delivery_sensitivity_as_per", "DeliverySensitivityAsPerRSC" },
        { "udo_dop", "DeliveryDateAsPerPO" },
        { "ucrd_dop", "CustomerRequiredDelivery" },
        { "ucdd_dop", "CustomerDelightDate" },
        { "ursc_dop", "RSCDelightDate" },
        { "uadd_dop", "AchievedDeliveryDate" },
        { "submitters_request_pd", "SubmittersRequest" },
        { "remarks_ot10", "Remarks" },
        { "k__uot_sold_party_dp", "CustomerID" },
        //{ "k__customer_contacts_dp1", "CustomerContactID" },
        { "k__uot_ship_to_party_dp", "EndUserID" },
        //{ "k__customer_contacts_dp3", "EndUserContactID" },
        //{ "k__customer_contacts_dp2", "CustomerContactID2" },
        //{ "k__customer_contacts_dp4", "EndUserContactID2" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" }
    };

    // Hardcoded column mapping for SparesOrderTransmittalLineItem table
    private static readonly Dictionary<string, string> SparesOrderTransmittalLineItemMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "SparesOrderTransmittalLineItemId" },
            { "record_id", "SparesOrderTransmittalID" },
            { "uuu_file_repo_name", "Name" },
            { "mandatory_documents_new_pd", "MandatoryDocuments" },
            { "manddatory_comments", "Comments" },
            { "group_type", "IsGroupType" },
            { "k__creator_id", "PrimaveraCreatedId" },
            { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
            { "uuu_li_last_update_date", "UpdatedAt" }
    };


    // Hardcoded column mapping for OrderReceiptAcknowledgement table
    private static readonly Dictionary<string, string> OrderReceiptAcknowledgementMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "ora_sec_sdt250", "SteamEndCasing" },
        { "ora_eec_sdt250", "ExhaustEndCasing" },
        { "ora_nc_sdt250", "NozzleChest" },
        { "ora_gbc_sdt250", "GBC" },
        { "ora_pedestal_sdt250", "Pedestal" },
        { "gear_box", "GearBox" },
        { "ora_bf_sdt250", "BladeForging" },
        { "ora_ppc_del_date_dop", "PPCDeliveryDate" },
        { "ora_marketing_fore_ynpd", "MarketingForecast" },
        { "ora_mfdate_dop", "MarketingForecastDate" },
        { "ora_riskdes_mtb4000", "RiskDescription" },
        { "ora_riskcons_mtb4000", "RiskConsequences" },
        { "agreed_sms", "AgreedSMS" },
        { "risk_of_ld", "RiskOfLD" },
        { "exesum_remarks_mtb4000", "InitiatorRemarks" },
        { "exesum_remarks2_mtb4000", "PPCRemarks" },
        { "exesum_remarks3_mtb4000", "PMRemarks" },
        { "exesum_remarks4_mtb4000", "AcknowledgementRemarks" },
        //{ "k__pp_select_frame_dpk", "SelectFrame" },
        { "ora_framenew_ynpd", "FrameNew" },
        { "ora_sel_lead_pd", "SelectLeadTime" },
        { "ora_sdt_w_rm_da", "TurbineDeliveryWithAvailableRMWeeks" },
        { "ora_sdt_wo_rm_da", "TurbineDeliveryWithProcurementWeeks" },
        { "ora_sdt_wo_rm_new_da", "TurbineDeliveryNewDevelopmentWeeks" },
        { "pp_lead_time_cc_da", "LeadTimeAsPerCCWeeks" },
        { "ora_rf_sdt250", "RotorForging" },
        { "uuu_creation_date", "CreatedAt" },
        { "k__creator_id", "CreatedBy" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "k__uuu_record_last_update_user", "UpdatedBy" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "id", "OrderReceiptAcknowledgementId" },
        { "k__pp_select_frame_dpk", "FrameMasterID" },
        { "project_id", "ProjectID" },
        { "ordertransmittalid", "OrderTransmittalID" }
    };

    private static readonly Dictionary<string, string> AuditActionMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "AuditActionId" },
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "ia_correction_done_df", "CorrectionDone" },
        { "ia_actual_corr_com_date_df", "ActualCorrectionCompletionDate" },
        { "ia_verification_decision_df", "VerificationDecision" },
        { "ia_veri_deci_date_df", "VerificationDecisionDate" },
        { "ia_reason_for_ver_dec_df", "ReasonForVerificationDecision" },
        { "ia_root_cause_analysis_df", "RootCauseAnalysis" },
        { "ia_actual_cor_act_com_date_df", "ActualCorrectiveActionCompletionDate" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "k__creator_id", "PrimaveraCreatedId" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user_name", "UpdatedName" },
        { "k__uuu_record_last_update_user", "PrimaveraUpdatedId" },
        { "k__ia_sel_obs_blip", "InternalAudit_ObservationDetail" },
        { "k__ia_sel_audit_rec_bpp", "InternalAuditId" },
        { "project_id", "ProjectID" },
        { "ia_corrective_action_tak_df", "CorrectiveActionTaken" }
    };

    private static readonly Dictionary<string, string> RCCAMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "RCCAId" },
            { "sqdcm_rcca_required_pd", "IsRCCARequired" },
            { "rcca_rcca_lead", "RCCALead" },
            { "rcca_lead_department", "LeadDepartment" },
            { "vendor_select_dp", "SelectSupplier" },
            { "rcca_sec_department_pd", "SecondaryDepartment" },
            { "rcca_why1_sdt250", "Why1" },
            { "rcca_why2_sdt250", "Why2" },
            { "rcca_why3_sdt250", "Why3" },
            { "rcca_why4_sdt250", "Why4" },
            { "rcca_why5_sdt250", "Why5" },
            { "rcca_root_cause", "RootCause" },
            { "rcca_rca_satisfactory", "RCASatisfactory" },
            { "rcca_remarks", "Remarks" },
            { "rcca_rca_ok", "IsRCAOk" },
            { "rcca_corrective_action", "CorrectiveAction" },
            { "rcca_proposed_date_of_imple", "ProposedDateOfImplementation" },
            { "rcca_imple_responsible", "ImplementationResponsible1" },
            { "rcca_imple_responsible1", "ImplementationResponsible2" },
            { "rcca_imple_responsible2", "ImplementationResponsible3" },
            { "rcca_imple_responsible3", "ImplementationResponsible4" },
            { "rcca_imple_responsible4", "ImplementationResponsible5" },
            { "rcca_imple_responsible5", "ImplementationResponsible6" },
            { "rcca_imple_responsible6", "ImplementationResponsible7" },
            { "rcca_imple_responsible7", "ImplementationResponsible8" },
            { "rcca_imple_responsible8", "ImplementationResponsible9" },
            { "rcca_imple_responsible9", "ImplementationResponsible10" },
            { "rcca_date_of_imple", "DateOfImplementation" },
            { "rcca_implem_corre_action", "ImplementedCorrectiveAction" },
            { "rcca_corrective_action_lead", "CorrectiveActionLead" },
            { "rcca_corre_action_lead_depa", "CorrectiveActionLeadDepartment" },
            { "rcca_pam_ynpd", "IsProActiveMeasuresRequired" },
            { "project_id", "ProjectID" },
            { "ordertransmittalid", "OrderTransmittalID" },
            { "status", "Status" },
            { "record_no", "RecordNo" },
            { "com_comp_rec_smn", "ComplaintRecordNo" },
            { "dev_rec_smn", "DeviationRequestNo" },
            { "creator_id", "CreatedName" },
            { "uuu_creation_date", "CreatedAt" },
            { "uuu_record_last_update_user", "UpdatedName" },
            { "uuu_record_last_update_date", "UpdatedAt" },
            { "k__creator_id", "PrimaveraCreatedId" },
            { "k__uuu_record_last_update_us", "PrimaveraUpdatedId" }
    };

    private static readonly Dictionary<string, string> RCCA_StandardLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "RCCA_StandardID" },
            { "record_id", "RCCAId" },
            { "notes", "Notes" },
            { "uuu_li_last_update_date", "UpdatedAt" }
    };

    private static readonly Dictionary<string, string> RCCA_SelectTeamMembersLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
            { "id", "RCCA_SelectTeamMemberID" },
            { "rcca_select_team_members_up", "SelectTeamMembers" },
            { "rcca_departments_df", "Departments" },
            { "record_id", "RCCAId" },
            { "uuu_li_last_update_date", "UpdatedAt" }
    };

    // Hardcoded column mapping for MonthlyProgressReport table
    private static readonly Dictionary<string, string> MonthlyProgressReportMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "record_no", "RecordNo" },
        { "status", "Status" },
        { "k__bp_creator_n", "BpCreator" },
        { "mpr_title_sdt120", "ProjectName" }, // Also maps to Title, handled in MatchColumns
        { "ec_m_proj_loc_sdt250", "ProjectLocation" },
        { "mpr_subtitle_sdt250", "Subtitle" },
        { "uot_ratiing_ia", "TurbineRating" },
        { "uot_nonstandard_pd", "SpecifyIfNonStandard" },
        { "mpr_scope_trans_pd", "ScopeOfTransportation" },
        { "mpr_scope_insurance_pd", "ScopeOfInsurance" },
        { "mpr_scope_services_pd", "ScopeOfServices" },
        { "mpr_services_type_pd", "ServicesType" },
        { "project_leader", "ProjectLeader" },
        { "udc_cur_amount_pd", "Currency" },
        { "mpr_ttl_pm_smn", "TTLsProjectManager" },
        { "uec_contact_no", "TTLsProjectManagerContactNo" },
        { "mpr_client_pm_smn", "ClientsProjectManager" },
        { "uec_contact_no1", "ClientsProjectManagerContactNo" },
        { "mpr_image1_ipk", "Image1" },
        { "mpr_image_desc_sdt120", "Image1Description" },
        { "mpr_image2_ipk", "Image2" },
        { "mpr_image_desc2_sdt120", "Image2Description" },
        { "mpr_image3_ipk", "Image3" },
        { "mpr_image_desc3_sdt120", "Image3Description" },
        { "mpr_image4_ipk", "Image4" },
        { "mpr_image_desc4_sdt120", "Image4Description" },
        { "mpr_image5_ipk", "Image5" },
        { "mpr_image_desc5_sdt120", "Image5Description" },
        { "mpr_image6_ipk", "Image6" },
        { "mpr_image_desc6_sdt120", "Image6Description" },
        { "mpr_image7_ipk", "Image7" },
        { "mpr_image_desc7_sdt120", "Image7Description" },
        { "mpr_image8_ipk", "Image8" },
        { "mpr_image_desc8_sdt120", "Image8Description" },
        { "mpr_image9_ipk", "Image9" },
        { "mpr_image_desc9_sdt120", "Image9Description" },
        { "mpr_image10_ipk", "Image10" },
        { "mpr_image_desc10_sdt120", "Image10Description" },
        { "mpr_image11_ipk", "Image11" },
        { "mpr_image_desc11_sdt120", "Image11Description" },
        { "mpr_image12_ipk", "Image12" },
        { "mpr_image_desc12_sdt120", "Image12Description" },
        { "mpr_image13_ipk", "Image13" },
        { "mpr_image_desc13_sdt120", "Image13Description" },
        { "mpr_image14_ipk", "Image14" },
        { "mpr_image_desc14_sdt120", "Image14Description" },
        { "mpr_image15_ipk", "Image15" },
        { "mpr_image_desc15_sdt120", "Image15Description" },
        { "mpr_image16_ipk", "Image16" },
        { "mpr_image_desc16_sdt120", "Image16Description" },
        { "project_id", "ProjectID" },
        { "ordertransmittalid", "OrderTransmittalID" },
        { "uuu_creation_date", "CreatedAt" },
        { "creator_id", "CreatedName" },
        { "uuu_record_last_update_date", "UpdatedAt" },
        { "uuu_record_last_update_user", "UpdatedName" },
        { "k__creator_id", "PrimeveraCreatedId" },
        { "k__uuu_record_last_update_user", "PrimeveraUpdatedId" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_MajorMilestoneLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_MajorMilestoneLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_MajorMilestoneID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_milestone_desc_sdt250", "MilestoneDescription" },
        { "mpr_date_comp_dop", "DateOfCompletion" },
        { "mpr_status_pd", "Status" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_ScopeOfSupplyLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_ScopeOfSupplyLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_ScopeOfSupplyID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_sr_no_da", "No" },
        { "mpr_type_pd", "Type" },
        { "mpr_item_desc_sdt250", "ItemDescription" },
        { "mpr_unit_of_mat_pd", "UnitOfMaterial" },
        { "mpr_quantity_da", "Quantity" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_ProcurementProgressofBoughtOutsItemLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_ProcurementProgressofBoughtOutsItemLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_ProcurementProgressofBoughtOutsItemID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_sr_no_da", "No" },
        { "mpr_type_pp_pd", "Type" },
        { "mpr_item_desc_sdt250", "ItemDescription" },
        { "mpr_po_date_dop", "PurchaseOrderDate" },
        { "mpr_expect_del_dop", "ExpectedDeliveryDate" },
        { "vendor_master_vendor", "Vendor" },
        { "mpr_current_stat_sdt250", "CurrentStatus" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_TurbineManufacturingProgressLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_TurbineManufacturingProgressLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_TurbineManufacturingProgressID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_item_desc_sdt250", "ItemDescription" },
        { "mpr_current_stat_sdt250", "CurrentStatus" },
        { "mpr_expec_date_com_dop", "ExpectedDateOfCompletion" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_InspectionDispatchPlanLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_InspectionDispatchPlanLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_InspectionDispatchPlanID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_item_desc_sdt250", "ItemDescription" },
        { "mpr_point_witness_sdt250", "PointOfWitness" },
        { "mpr_place_inspec_sdt250", "PlaceOfInspection" },
        { "mpr_inspec_date_dop", "InspectionDate" },
        { "mpr_dispatch_date_dop", "DispatchDate" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_CashInFlowPlanLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_CashInFlowPlanLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_CashInFlowPlanID" },
        { "k__ci_sel_ini_cf_plan_dpk", "InitialCashPlanId" },
        { "k__mpr_sel_month_lidp", "InitialCashFlowPlanId" },
        { "uec_payment_terms", "PaymentTerms" },
        { "ci_mil_per_da", "MilestonePercentage" },
        { "ci_tot_net_payable_da", "TotalNetPayableAmount" },
        { "ci_planned_monthh_dp", "PlannedMonth" },
        { "mpr_act_date_rec_dop", "ActualDateOfReceipt" },
        { "record_id", "MonthlyProgressReportID" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_LookAheadTaskforNext30DaysLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_LookAheadTaskforNext30DaysLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_LookAheadTaskforNext30DaysID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_task_desc_sdt250", "TaskDescription" },
        { "mpr_date_comp_dop", "DateOfCompletion" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_EngineeringProgressLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_EngineeringProgressLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_EngineeringProgressID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_area_pd", "Area" },
        { "remarks_ot1", "Remarks" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_InputsRequiredFromCustomerLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_InputsRequiredFromCustomerLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_InputsRequiredFromCustomerID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_areas_concern_mtb4000", "InputsRequiredFromCustomer" }
    };

    // Hardcoded column mapping for MonthlyProgressReport_DBOSummarizeSheetLI table
    private static readonly Dictionary<string, string> MonthlyProgressReport_DBOSummarizeSheetLIMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "MPR_DBOSummarizeSheetID" },
        { "record_id", "MonthlyProgressReportID" },
        { "mpr_type_pd", "Type" },
        { "mpr_scope_rb", "IsScopeOfSupply" },
        { "mpr_procurement_rb", "IsProcurementProgressBOI" },
        { "mpr_inspection_rb", "IsInspectionDispatchPlan" },
        { "mpr_item_desc_sdt250", "ItemDescription" },
        { "mpr_unit_of_mat_pd", "UnitOfMaterial" },
        { "mpr_quantity_da", "Quantity" },
        { "vendor_master_vendor", "Vendor" },
        { "mpr_po_no_sdt50", "PurchaseOrderNo" },
        { "mpr_po_date_dop", "PurchaseOrderDate" },
        { "mpr_point_witness_sdt250", "PointOfWitness" },
        { "mpr_place_inspec_sdt250", "PlaceOfInspection" },
        { "mpr_inspec_date_dop", "InspectionDate" },
        { "mpr_dispatch_date_dop", "DispatchDate" },
        { "mpr_current_stat_pd", "CurrentStatus" }
    };

    // Hardcoded column mapping for ApprovalLog table
    private static readonly Dictionary<string, string> ApprovalLogColumnMapping = new(StringComparer.OrdinalIgnoreCase)
{
     { "task_id", "ApprovalLogID" },
     { "process_source_id", "CompanyId" },
     { "process_project_id", "ProjectID" },
     { "task_round", "SequenceNumber" },
     { "task_name", "StepName" },
     { "assignee_id", "ApproverUserPrimaveraId" },
     { "action_name", "ActionName" },
     { "task_status", "StatusID" },
     { "source_modelname", "ApprovalUrl" },
     { "task_start_date", "CreatedAt" },
     { "task_end_date", "CompletionDate" }
};

    // Hardcoded column mapping for AuditLog table changes
    private static readonly Dictionary<string, string> AuditLogMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "id", "AuditID" },
        { "projectid", "ProjectID" },
        { "fieldname", "Field" },
        { "oldvalue", "OldValue" },
        { "newvalue", "NewValue" },
        { "eventname", "Event" },
        { "attchcnt", "Attachment" },
        { "proxyuserid", "ProxyUser" },
        { "eventtime", "CreatedAt" },
        { "sourceid", "SourceId" },
        { "sourcetype", "SourceType" }
    };
    public async Task<UploadResponse> MigrateExcelToSqlServerAsync(
        string connectionString,
        string schemaName,
        string tableName,
        DataTable excelData,
        string? attachmentRecordType = null,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        if (excelData == null || excelData.Rows.Count == 0)
        {
            response.ErrorMessages.Add("Excel file contains no data.");
            return response;
        }

        // Check if table name starts with "OrderTransmittal" OR is "Payment_Supply" OR "LiquidatedDamage" OR "Payment_ENC" - migrate to all matching tables
        if (string.Equals(tableName, "OrderTransmittal_Notes", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToOrderTransmittalNotesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        if (tableName.StartsWith("OrderTransmittal", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "Payment_Supply", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "payment_supply", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "LiquidatedDamage", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "liquidated_damage", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "Payment_ENC", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "payment_enc", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToOrderTransmittalTablesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        if (string.Equals(tableName, "OTBankGuarantee", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateOrderTransmittalLineItemBankGuaranteeAsync(
                connectionString,
                schemaName,
                excelData,
                attachmentRecordType,
                cancellationToken);
        }

        if (string.Equals(tableName, "BankGuaranteeSpares", StringComparison.OrdinalIgnoreCase))
        {
            await using var bankGuaranteeSparesConnection = new SqlConnection(connectionString);
            await bankGuaranteeSparesConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(bankGuaranteeSparesConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "DeviationRequest", StringComparison.OrdinalIgnoreCase))
        {
            await using var deviationRequestConnection = new SqlConnection(connectionString);
            await deviationRequestConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(deviationRequestConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "DesignRecommendation", StringComparison.OrdinalIgnoreCase))
        {
            await using var designRecommendationConnection = new SqlConnection(connectionString);
            await designRecommendationConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(designRecommendationConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "FrameMaster", StringComparison.OrdinalIgnoreCase))
        {
            await using var frameMasterConnection = new SqlConnection(connectionString);
            await frameMasterConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(frameMasterConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "BankMaster", StringComparison.OrdinalIgnoreCase))
        {
            await using var bankMasterConnection = new SqlConnection(connectionString);
            await bankMasterConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(bankMasterConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is UserList table - use single table migration
        if (string.Equals(tableName, "UserList", StringComparison.OrdinalIgnoreCase))
        {
            await using var userListConnection = new SqlConnection(connectionString);
            await userListConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(userListConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is MonthlyActualCollection table - use single table migration
        if (string.Equals(tableName, "MonthlyActualCollection", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("MonthlyActualCollection", StringComparison.OrdinalIgnoreCase))
        {
            await using var macConnection = new SqlConnection(connectionString);
            await macConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(macConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is InitialCashFlowPlan table - use single table migration
        if (string.Equals(tableName, "InitialCashFlowPlan", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("InitialCashFlowPlan", StringComparison.OrdinalIgnoreCase))
        {
            await using var icfpConnection = new SqlConnection(connectionString);
            await icfpConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(icfpConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is MonthlyPlanning table - use single table migration
        if (string.Equals(tableName, "MonthlyPlanning", StringComparison.OrdinalIgnoreCase))
        {
            await using var mpConnection = new SqlConnection(connectionString);
            await mpConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(mpConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is MonthlyPlanningLineItem table - use single table migration
        if (string.Equals(tableName, "MonthlyPlanningLineItem", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "MonthlyPlanning_LineItem", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("MonthlyPlanningLineItem", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("MonthlyPlanning_LineItem", StringComparison.OrdinalIgnoreCase))
        {
            await using var mpliConnection = new SqlConnection(connectionString);
            await mpliConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(mpliConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is MonthlyActualCollectionPlanned table - use single table migration
        if (string.Equals(tableName, "MonthlyActualCollectionPlanned", StringComparison.OrdinalIgnoreCase))
        {
            await using var macpConnection = new SqlConnection(connectionString);
            await macpConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(macpConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is MonthlyActualUnplannedCollection table - use single table migration
        if (string.Equals(tableName, "MonthlyActualUnplannedCollection", StringComparison.OrdinalIgnoreCase))
        {
            await using var maupConnection = new SqlConnection(connectionString);
            await maupConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(maupConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is SparesOrderTransmittal table - use single table migration
        if (string.Equals(tableName, "SparesOrderTransmittal", StringComparison.OrdinalIgnoreCase))
        {
            await using var sotConnection = new SqlConnection(connectionString);
            await sotConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(sotConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "SparesOrderTransmittalLineItem", StringComparison.OrdinalIgnoreCase))
        {
            await using var sotliConnection = new SqlConnection(connectionString);
            await sotliConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(sotliConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }



        if (string.Equals(tableName, "OrderReceiptAcknowledgement", StringComparison.OrdinalIgnoreCase))
        {
            await using var oraConnection = new SqlConnection(connectionString);
            await oraConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(oraConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "AuditAction", StringComparison.OrdinalIgnoreCase))
        {
            await using var aaConnection = new SqlConnection(connectionString);
            await aaConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(aaConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if table name is "RCCA_StandardLI" - specialized LI migration
        if (string.Equals(tableName, "RCCA_StandardLI", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToRCCALineItemsAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        // Check if table name is "RCCA_SelectTeamMembersLI" - specialized LI migration
        if (string.Equals(tableName, "RCCA_SelectTeamMembersLI", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToRCCALineItemsAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        if (string.Equals(tableName, "RCCA", StringComparison.OrdinalIgnoreCase))
        {
            await using var rccaConnection = new SqlConnection(connectionString);
            await rccaConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(rccaConnection, schemaName, tableName, excelData, null, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyProgressReport", StringComparison.OrdinalIgnoreCase))
        {
            await using var mprConnection = new SqlConnection(connectionString);
            await mprConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(mprConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (tableName.StartsWith("MonthlyProgressReport_", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToMonthlyProgressReportLineItemsAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        // Check if this is SpecificationRelease table - use single table migration
        if (string.Equals(tableName, "SpecificationRelease", StringComparison.OrdinalIgnoreCase))
        {
            await using var srConnection = new SqlConnection(connectionString);
            await srConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(srConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }
        // Check if this is ApprovalLog table - use single table migration
        if (string.Equals(tableName, "ApprovalLog", StringComparison.OrdinalIgnoreCase))
        {
            await using var srConnection = new SqlConnection(connectionString);
            await srConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(srConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "AuditLog", StringComparison.OrdinalIgnoreCase))
        {
            await using var auditLogConnection = new SqlConnection(connectionString);
            await auditLogConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(auditLogConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "Risks", StringComparison.OrdinalIgnoreCase))
        {
            await using var risksConnection = new SqlConnection(connectionString);
            await risksConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(risksConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "Risk_MitigationPlan", StringComparison.OrdinalIgnoreCase))
        {
            await using var riskMitigationPlanConnection = new SqlConnection(connectionString);
            await riskMitigationPlanConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(riskMitigationPlanConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "PositiveRecall", StringComparison.OrdinalIgnoreCase))
        {
            await using var positiveRecallConnection = new SqlConnection(connectionString);
            await positiveRecallConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(positiveRecallConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyCashInflow", StringComparison.OrdinalIgnoreCase))
        {
            await using var monthlyCashInflowConnection = new SqlConnection(connectionString);
            await monthlyCashInflowConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(monthlyCashInflowConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyCashInflow_CustomerCare", StringComparison.OrdinalIgnoreCase))
        {
            await using var monthlyCashInflowCustomerCareConnection = new SqlConnection(connectionString);
            await monthlyCashInflowCustomerCareConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(monthlyCashInflowCustomerCareConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyCashInflow_MarketingAdvance", StringComparison.OrdinalIgnoreCase))
        {
            await using var monthlyCashInflowMarketingAdvanceConnection = new SqlConnection(connectionString);
            await monthlyCashInflowMarketingAdvanceConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(monthlyCashInflowMarketingAdvanceConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyCashInflow_ICDue", StringComparison.OrdinalIgnoreCase))
        {
            await using var monthlyCashInflowICDueConnection = new SqlConnection(connectionString);
            await monthlyCashInflowICDueConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(monthlyCashInflowICDueConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "MonthlyCashInflow_ICNotDue", StringComparison.OrdinalIgnoreCase))
        {
            await using var monthlyCashInflowICNotDueConnection = new SqlConnection(connectionString);
            await monthlyCashInflowICNotDueConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(monthlyCashInflowICNotDueConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "HandoverProtocol", StringComparison.OrdinalIgnoreCase))
        {
            await using var handoverProtocolConnection = new SqlConnection(connectionString);
            await handoverProtocolConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(handoverProtocolConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "HandoverProtocolLineItem", StringComparison.OrdinalIgnoreCase))
        {
            await using var handoverProtocolLineItemConnection = new SqlConnection(connectionString);
            await handoverProtocolLineItemConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(handoverProtocolLineItemConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "InternalAudit", StringComparison.OrdinalIgnoreCase))
        {
            await using var internalAuditConnection = new SqlConnection(connectionString);
            await internalAuditConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(internalAuditConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "InternalAudit_Notes", StringComparison.OrdinalIgnoreCase))
        {
            await using var internalAuditNotesConnection = new SqlConnection(connectionString);
            await internalAuditNotesConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(internalAuditNotesConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "InternalAudit_ObservationDetails", StringComparison.OrdinalIgnoreCase))
        {
            await using var internalAuditObservationDetailsConnection = new SqlConnection(connectionString);
            await internalAuditObservationDetailsConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(internalAuditObservationDetailsConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        if (string.Equals(tableName, "OTRisk", StringComparison.OrdinalIgnoreCase))
        {
            await using var otRiskConnection = new SqlConnection(connectionString);
            await otRiskConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(otRiskConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }


        // Check if table name starts with "MechanicalDBO" - migrate to all matching tables
        if (tableName.StartsWith("MechanicalDBO", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToMechanicalDBOTablesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        // Check if this is BPComments table - use single table migration
        if (string.Equals(tableName, "BPComments", StringComparison.OrdinalIgnoreCase))
        {
            await using var bpCommentsConnection = new SqlConnection(connectionString);
            await bpCommentsConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(bpCommentsConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if this is BPAttachments table - use single table migration with attachmentRecordType
        if (string.Equals(tableName, "BPAttachments", StringComparison.OrdinalIgnoreCase))
        {
            await using var bpAttachmentsConnection = new SqlConnection(connectionString);
            await bpAttachmentsConnection.OpenAsync(cancellationToken);
            return await MigrateToSingleTableAsync(bpAttachmentsConnection, schemaName, tableName, excelData, attachmentRecordType, cancellationToken);
        }

        // Check if table name starts with "ElectricalInstrumentationDBO" - migrate to all matching tables
        if (tableName.StartsWith("ElectricalInstrumentationDBO", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToElectricalInstrumentationDBOTablesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        // Check if table name starts with "Turbine" - migrate to all matching tables
        if (tableName.StartsWith("Turbine", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToTurbineTablesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }

        if (tableName.StartsWith("MinutesOfMeeting", StringComparison.OrdinalIgnoreCase) || 
            tableName.StartsWith("MOM", StringComparison.OrdinalIgnoreCase))
        {
            return await MigrateToMinutesOfMeetingTablesAsync(connectionString, schemaName, tableName, excelData, cancellationToken);
        }




        // For single table migration (existing logic)
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        var transaction = connection.BeginTransaction();
        var tempTableName = $"#TMP_{Guid.NewGuid():N}";

        try
        {
            // Step 1: Get target table metadata
            var tableMetadata = await GetTableMetadataAsync(connection, transaction, schemaName, tableName, cancellationToken);

            if (tableMetadata.Count == 0)
            {
                response.ErrorMessages.Add($"Table '{schemaName}.{tableName}' not found or has no columns.");
                transaction.Rollback();
                return response;
            }

            // Step 2: Match Excel columns to SQL columns
            var columnMappings = MatchColumns(excelData, tableMetadata, tableName, attachmentRecordType);

            if (columnMappings.Count == 0)
            {
                response.ErrorMessages.Add("No matching columns found between Excel and SQL table.");
                transaction.Rollback();
                return response;
            }

            // Step 3: Check for identity column
            var identityColumn = tableMetadata.FirstOrDefault(m => m.IsIdentity);
            var hasIdentityInExcel = identityColumn != null &&
                                    columnMappings.Any(m => m.SqlColumnName.Equals(identityColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

            // Step 4: Create temp table
            await CreateTempTableAsync(connection, transaction, tempTableName, tableMetadata, cancellationToken);

            // Step 5: Prepare DataTable with only matched columns
            var (mappedDataTable, rowErrors) = await PrepareMappedDataTableAsync(connection, transaction, excelData, columnMappings, tableName, schemaName, tableMetadata, cancellationToken);

            // Add row errors to response
            response.RowErrors.AddRange(rowErrors);

            // Step 6: Bulk copy to temp table
            var rowsCopiedToTemp = await BulkCopyToTempTableAsync(
                connection,
                transaction,
                tempTableName,
                mappedDataTable,
                columnMappings,
                hasIdentityInExcel,
                cancellationToken);

            // Step 7: Get primary key columns
            var primaryKeyColumns = tableMetadata.Where(m => m.IsPrimaryKey).Select(m => m.ColumnName).ToList();

            // Step 8: Upsert from temp table to target table using MERGE
            var (rowsInserted, rowsUpdated) = await MergeFromTempToTargetAsync(
                connection,
                transaction,
                schemaName,
                tableName,
                tempTableName,
                columnMappings,
                primaryKeyColumns,
                identityColumn,
                hasIdentityInExcel,
                cancellationToken);

            transaction.Commit();

            response.Success = rowErrors.Count == 0;
            response.RowsInserted = rowsInserted;
            response.RowsUpdated = rowsUpdated;
            response.RowsFailed = rowErrors.Count;

            var totalProcessed = rowsInserted + rowsUpdated;
            if (totalProcessed > 0)
            {
                response.Message = $"Successfully processed {totalProcessed} row(s): {rowsInserted} inserted, {rowsUpdated} updated.";
            }

            if (rowErrors.Count > 0)
            {
                response.ErrorMessages.Add($"{rowErrors.Count} row(s) failed during data preparation. See RowErrors for details.");
            }

        }
        catch (Exception ex)
        {
            try
            {
                transaction.Rollback();
            }
            catch
            {
                // Ignore rollback errors if transaction already completed
            }
            
            response.ErrorMessages.Add($"Error during migration: {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }
        finally
        {
            // Clean up temp table - pass null for transaction as it's already committed or rolled back
            try
            {
                await DropTempTableAsync(connection, null, tempTableName, cancellationToken);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        return response;
    }

    private async Task<UploadResponse> MigrateToOrderTransmittalTablesAsync(
        string connectionString,
        string schemaName,
        string tableNamePrefix,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();
        var allTableResults = new List<(string tableName, int inserted, int updated, int failed, List<string> errors)>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        try
        {
            // Step 1: Find all tables matching the criteria
            var matchingTables = new List<string>();

            // If user specifically asked for "Payment_Supply", only migrate that table
            if (string.Equals(tableNamePrefix, "Payment_Supply", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableNamePrefix, "payment_supply", StringComparison.OrdinalIgnoreCase))
            {
                var paymentSupplyTables = await GetTablesWithPrefixAsync(connection, schemaName, "payment_supply", cancellationToken);
                matchingTables.AddRange(paymentSupplyTables);
            }
            else if (string.Equals(tableNamePrefix, "Payment_ENC", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableNamePrefix, "payment_enc", StringComparison.OrdinalIgnoreCase))
            {
                var paymentEncTables = await GetTablesWithPrefixAsync(connection, schemaName, "payment_enc", cancellationToken);
                matchingTables.AddRange(paymentEncTables);
            }
            else if (string.Equals(tableNamePrefix, "LiquidatedDamage", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableNamePrefix, "liquidated_damage", StringComparison.OrdinalIgnoreCase))
            {
                var liquidatedDamageTables = await GetTablesWithPrefixAsync(connection, schemaName, "LiquidatedDamage", cancellationToken);
                if (liquidatedDamageTables.Count == 0)
                {
                    liquidatedDamageTables = await GetTablesWithPrefixAsync(connection, schemaName, "liquidated_damage", cancellationToken);
                }
                matchingTables.AddRange(liquidatedDamageTables);
            }
            else
            {
                // Default behavior: Migrate OrderTransmittal ONLY
                var otTables = await GetTablesWithPrefixAsync(connection, schemaName, "OrderTransmittal", cancellationToken);
                matchingTables.AddRange(otTables);
                // Do NOT include payment_supply unless explicitly requested
            }

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No tables found with prefix 'OrderTransmittal' in schema '{schemaName}'.");
                return response;
            }

            // Step 2: Sort tables to ensure parent table is migrated first
            // Parent table is "OrderTransmittal" (exact match), child tables have underscores
            var parentTable = matchingTables.FirstOrDefault(t =>
                string.Equals(t, "OrderTransmittal", StringComparison.OrdinalIgnoreCase));
            var childTables = matchingTables.Where(t =>
                !string.Equals(t, "OrderTransmittal", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t).ToList();

            // Build ordered list: parent first, then children
            var orderedTables = new List<string>();
            if (parentTable != null)
            {
                orderedTables.Add(parentTable);
            }
            orderedTables.AddRange(childTables);

            // Step 3: Migrate Excel data to each matching table in order
            foreach (var targetTable in orderedTables)
            {
                DataTable tableSpecificData = excelData;
                var uuuTabIdColumn = excelData.Columns.Cast<DataColumn>()
                    .FirstOrDefault(c => c.ColumnName.Trim().Equals("uuu_tab_id", StringComparison.OrdinalIgnoreCase))?.ColumnName;

                bool isPaymentSupplyData = uuuTabIdColumn != null && 
                                           excelData.AsEnumerable().Any(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "11" || v == "11.0"; });
                bool isLiquidatedDamageData = uuuTabIdColumn != null && 
                                              excelData.AsEnumerable().Any(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "4" || v == "4.0"; });
                bool isPaymentEncData = uuuTabIdColumn != null && 
                                        excelData.AsEnumerable().Any(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "12" || v == "12.0"; });
                
                // If the Excel data is for payment_supply (uuu_tab_id=11), LiquidatedDamage (4), or payment_enc (12), SKIP the main OrderTransmittal table
                if ((isPaymentSupplyData || isLiquidatedDamageData || isPaymentEncData) && string.Equals(targetTable, "OrderTransmittal", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Special row filtering for payment_supply child table based on uuu_tab_id
                if (string.Equals(targetTable, "payment_supply", StringComparison.OrdinalIgnoreCase))
                {
                    if (uuuTabIdColumn != null)
                    {
                        var filteredRows = excelData.AsEnumerable()
                            .Where(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "11" || v == "11.0"; })
                            .ToList();

                        if (filteredRows.Count > 0)
                        {
                            tableSpecificData = filteredRows.CopyToDataTable();
                        }
                        else
                        {
                            continue; // No matching rows for payment_supply
                        }
                    }
                    else
                    {
                        continue; // Column missing
                    }
                }

                // Special row filtering for LiquidatedDamage child table based on uuu_tab_id
                if (string.Equals(targetTable, "LiquidatedDamage", StringComparison.OrdinalIgnoreCase) || string.Equals(targetTable, "liquidated_damage", StringComparison.OrdinalIgnoreCase))
                {
                    if (uuuTabIdColumn != null)
                    {
                        var filteredRows = excelData.AsEnumerable()
                            .Where(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "4" || v == "4.0"; })
                            .ToList();

                        if (filteredRows.Count > 0)
                        {
                            tableSpecificData = filteredRows.CopyToDataTable();
                        }
                        else
                        {
                            continue; // No matching rows for LiquidatedDamage
                        }
                    }
                    else
                    {
                        continue; // Column missing
                    }
                }

                // Special row filtering for payment_enc child table based on uuu_tab_id
                if (string.Equals(targetTable, "payment_enc", StringComparison.OrdinalIgnoreCase))
                {
                    if (uuuTabIdColumn != null)
                    {
                        var filteredRows = excelData.AsEnumerable()
                            .Where(r => { var v = r[uuuTabIdColumn]?.ToString()?.Trim(); return v == "12" || v == "12.0"; })
                            .ToList();

                        if (filteredRows.Count > 0)
                        {
                            tableSpecificData = filteredRows.CopyToDataTable();
                        }
                        else
                        {
                            continue; // No matching rows for payment_enc
                        }
                    }
                    else
                    {
                        continue; // Column missing
                    }
                }

                var tableResponse = await MigrateToSingleTableAsync(
                    connection,
                    schemaName,
                    targetTable,
                    tableSpecificData,
                    null,
                    cancellationToken);

                allTableResults.Add((
                    targetTable,
                    tableResponse.RowsInserted,
                    tableResponse.RowsUpdated,
                    tableResponse.RowsFailed,
                    tableResponse.ErrorMessages.ToList()
                ));

                // Aggregate row errors
                response.RowErrors.AddRange(tableResponse.RowErrors);
            }

            // Step 3: Aggregate results
            response.Success = allTableResults.All(r => r.errors.Count == 0) && response.RowErrors.Count == 0;
            response.RowsInserted = allTableResults.Sum(r => r.inserted);
            response.RowsUpdated = allTableResults.Sum(r => r.updated);
            response.RowsFailed = allTableResults.Sum(r => r.failed) + response.RowErrors.Count;

            // Build summary message
            var successCount = allTableResults.Count(r => r.errors.Count == 0);
            var totalTables = allTableResults.Count;
            var totalProcessed = response.RowsInserted + response.RowsUpdated;

            if (totalProcessed > 0)
            {
                response.Message = $"Migrated to {totalTables} table(s): {successCount} succeeded. " +
                                 $"Total: {totalProcessed} row(s) processed ({response.RowsInserted} inserted, {response.RowsUpdated} updated).";
            }

            // Add per-table error messages
            foreach (var result in allTableResults.Where(r => r.errors.Count > 0))
            {
                response.ErrorMessages.Add($"Table '{result.tableName}': {string.Join("; ", result.errors)}");
            }

            if (response.RowErrors.Count > 0)
            {
                response.ErrorMessages.Add($"{response.RowErrors.Count} row(s) failed during data preparation. See RowErrors for details.");
            }
        }
        catch (Exception ex)
        {
            response.ErrorMessages.Add($"Error during OrderTransmittal migration: {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }

        return response;
    }

    private async Task<UploadResponse> MigrateToTurbineTablesAsync(
        string connectionString,
        string schemaName,
        string tableNamePrefix,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();
        var allTableResults = new List<(string tableName, int inserted, int updated, int failed, List<string> errors)>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        try
        {
            // Step 1: Find all tables starting with "Turbine" in the schema
            var matchingTables = await GetTablesWithPrefixAsync(connection, schemaName, "Turbine", cancellationToken);

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No tables found with prefix 'Turbine' in schema '{schemaName}'.");
                return response;
            }

            // Step 2: Sort tables to ensure parent table is migrated first
            // Parent table is "Turbine" (exact match), child tables have underscores
            var parentTable = matchingTables.FirstOrDefault(t =>
                string.Equals(t, "Turbine", StringComparison.OrdinalIgnoreCase));
            var childTables = matchingTables.Where(t =>
                !string.Equals(t, "Turbine", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t).ToList();

            // Build ordered list: parent first, then children
            var orderedTables = new List<string>();
            if (parentTable != null)
            {
                orderedTables.Add(parentTable);
            }
            orderedTables.AddRange(childTables);

            // Step 3: Migrate Excel data to each matching table in order
            foreach (var targetTable in orderedTables)
            {
                var tableResponse = await MigrateToSingleTableAsync(
                    connection,
                    schemaName,
                    targetTable,
                    excelData,
                    null,
                    cancellationToken);

                allTableResults.Add((
                    targetTable,
                    tableResponse.RowsInserted,
                    tableResponse.RowsUpdated,
                    tableResponse.RowsFailed,
                    tableResponse.ErrorMessages.ToList()
                ));

                // Aggregate row errors
                response.RowErrors.AddRange(tableResponse.RowErrors);
            }

            // Step 4: Aggregate results
            response.Success = allTableResults.All(r => r.errors.Count == 0) && response.RowErrors.Count == 0;
            response.RowsInserted = allTableResults.Sum(r => r.inserted);
            response.RowsUpdated = allTableResults.Sum(r => r.updated);
            response.RowsFailed = allTableResults.Sum(r => r.failed) + response.RowErrors.Count;

            // Build summary message
            var successCount = allTableResults.Count(r => r.errors.Count == 0);
            var totalTables = allTableResults.Count;
            var totalProcessed = response.RowsInserted + response.RowsUpdated;

            if (totalProcessed > 0)
            {
                response.Message = $"Migrated to {totalTables} table(s): {successCount} succeeded. " +
                                 $"Total: {totalProcessed} row(s) processed ({response.RowsInserted} inserted, {response.RowsUpdated} updated).";
            }

            // Add per-table error messages
            foreach (var result in allTableResults.Where(r => r.errors.Count > 0))
            {
                response.ErrorMessages.Add($"Table '{result.tableName}': {string.Join("; ", result.errors)}");
            }

            if (response.RowErrors.Count > 0)
            {
                response.ErrorMessages.Add($"{response.RowErrors.Count} row(s) failed during data preparation. See RowErrors for details.");
            }
        }
        catch (Exception ex)
        {
            response.ErrorMessages.Add($"Error during Turbine migration: {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }

        return response;
    }

    private async Task<UploadResponse> MigrateToMechanicalDBOTablesAsync(
        string connectionString,
        string schemaName,
        string tableNamePrefix,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();
        var allTableResults = new List<(string tableName, int inserted, int updated, int failed, List<string> errors)>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        try
        {
            // Step 1: Find all tables starting with "MechanicalDBO" in the schema
            var matchingTables = await GetTablesWithPrefixAsync(connection, schemaName, "MechanicalDBO", cancellationToken);

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No tables found with prefix 'MechanicalDBO' in schema '{schemaName}'.");
                return response;
            }

            // Step 2: Sort tables to ensure parent table is migrated first
            // Parent table is "MechanicalDBO" (exact match), child tables have underscores
            var parentTable = matchingTables.FirstOrDefault(t =>
                string.Equals(t, "MechanicalDBO", StringComparison.OrdinalIgnoreCase));
            var childTables = matchingTables.Where(t =>
                !string.Equals(t, "MechanicalDBO", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t).ToList();

            // Build ordered list: parent first, then children
            var orderedTables = new List<string>();
            if (parentTable != null)
            {
                orderedTables.Add(parentTable);
            }
            orderedTables.AddRange(childTables);

            // Step 3: Migrate Excel data to each matching table in order
            foreach (var targetTable in orderedTables)
            {
                var tableResponse = await MigrateToSingleTableAsync(
                    connection,
                    schemaName,
                    targetTable,
                    excelData,
                    null,
                    cancellationToken);

                allTableResults.Add((
                    targetTable,
                    tableResponse.RowsInserted,
                    tableResponse.RowsUpdated,
                    tableResponse.RowsFailed,
                    tableResponse.ErrorMessages.ToList()
                ));

                // Aggregate row errors
                response.RowErrors.AddRange(tableResponse.RowErrors);
            }

            // Step 4: Aggregate results
            response.Success = allTableResults.All(r => r.errors.Count == 0) && response.RowErrors.Count == 0;
            response.RowsInserted = allTableResults.Sum(r => r.inserted);
            response.RowsUpdated = allTableResults.Sum(r => r.updated);
            response.RowsFailed = allTableResults.Sum(r => r.failed) + response.RowErrors.Count;

            // Build summary message
            var successCount = allTableResults.Count(r => r.errors.Count == 0);
            var totalTables = allTableResults.Count;
            var totalProcessed = response.RowsInserted + response.RowsUpdated;

            if (totalProcessed > 0)
            {
                response.Message = $"Migrated to {totalTables} table(s): {successCount} succeeded. " +
                                 $"Total: {totalProcessed} row(s) processed ({response.RowsInserted} inserted, {response.RowsUpdated} updated).";
            }

            // Add per-table error messages
            foreach (var result in allTableResults.Where(r => r.errors.Count > 0))
            {
                response.ErrorMessages.Add($"Table '{result.tableName}': {string.Join("; ", result.errors)}");
            }

            if (response.RowErrors.Count > 0)
            {
                response.ErrorMessages.Add($"{response.RowErrors.Count} row(s) failed during data preparation. See RowErrors for details.");
            }
        }
        catch (Exception ex)
        {
            response.ErrorMessages.Add($"Error during MechanicalDBO migration: {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }

        return response;
    }

    private async Task<UploadResponse> MigrateToElectricalInstrumentationDBOTablesAsync(
        string connectionString,
        string schemaName,
        string tableNamePrefix,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();
        var allTableResults = new List<(string tableName, int inserted, int updated, int failed, List<string> errors)>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        try
        {
            // Step 1: Find matching tables.
            // - Exact child: migrate only that child table.
            // - Parent/prefix: migrate parent + child tables.
            List<string> matchingTables;
            if (tableNamePrefix.StartsWith("ElectricalInstrumentationDBO_", StringComparison.OrdinalIgnoreCase))
            {
                matchingTables = new List<string> { tableNamePrefix };
            }
            else
            {
                matchingTables = await GetTablesWithPrefixAsync(connection, schemaName, "ElectricalInstrumentationDBO", cancellationToken);
            }

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No tables found with prefix 'ElectricalInstrumentationDBO' in schema '{schemaName}'.");
                return response;
            }

            // Keep only existing tables and remove duplicates.
            var availableElectricalTables = await GetTablesWithPrefixAsync(connection, schemaName, "ElectricalInstrumentationDBO", cancellationToken);
            matchingTables = matchingTables
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(mt => availableElectricalTables.Any(at => string.Equals(at, mt, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No matching ElectricalInstrumentationDBO table found for '{tableNamePrefix}' in schema '{schemaName}'.");
                return response;
            }

            // Step 2: Sort tables to ensure parent table is migrated first
            // Parent table is "ElectricalInstrumentationDBO" (exact match), child tables have underscores
            var parentTable = matchingTables.FirstOrDefault(t =>
                string.Equals(t, "ElectricalInstrumentationDBO", StringComparison.OrdinalIgnoreCase));
            var childTables = matchingTables.Where(t =>
                !string.Equals(t, "ElectricalInstrumentationDBO", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t).ToList();

            // Build ordered list: parent first, then children
            var orderedTables = new List<string>();
            if (parentTable != null)
            {
                orderedTables.Add(parentTable);
            }
            orderedTables.AddRange(childTables);

            // Step 3: Migrate Excel data to each matching table in order
            foreach (var targetTable in orderedTables)
            {
                var tableResponse = await MigrateToSingleTableAsync(
                    connection,
                    schemaName,
                    targetTable,
                    excelData,
                    null,
                    cancellationToken,
                    "ElectricalInstrumentationDBO");

                allTableResults.Add((
                    targetTable,
                    tableResponse.RowsInserted,
                    tableResponse.RowsUpdated,
                    tableResponse.RowsFailed,
                    tableResponse.ErrorMessages.ToList()
                ));

                // Aggregate row errors
                response.RowErrors.AddRange(tableResponse.RowErrors);
            }

            // Step 4: Aggregate results
            response.Success = allTableResults.All(r => r.errors.Count == 0) && response.RowErrors.Count == 0;
            response.RowsInserted = allTableResults.Sum(r => r.inserted);
            response.RowsUpdated = allTableResults.Sum(r => r.updated);
            response.RowsFailed = allTableResults.Sum(r => r.failed) + response.RowErrors.Count;

            // Build summary message
            var successCount = allTableResults.Count(r => r.errors.Count == 0);
            var totalTables = allTableResults.Count;
            var totalProcessed = response.RowsInserted + response.RowsUpdated;

            if (totalProcessed > 0)
            {
                response.Message = $"Migrated to {totalTables} table(s): {successCount} succeeded. " +
                                 $"Total: {totalProcessed} row(s) processed ({response.RowsInserted} inserted, {response.RowsUpdated} updated).";
            }

            // Add per-table error messages
            foreach (var result in allTableResults.Where(r => r.errors.Count > 0))
            {
                response.ErrorMessages.Add($"Table '{result.tableName}': {string.Join("; ", result.errors)}");
            }

            if (response.RowErrors.Count > 0)
            {
                response.ErrorMessages.Add($"{response.RowErrors.Count} row(s) failed during data preparation. See RowErrors for details.");
            }
        }
        catch (Exception ex)
        {
            response.ErrorMessages.Add($"Error during ElectricalInstrumentationDBO migration: {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }

        return response;
    }



    private async Task<UploadResponse> MigrateToMinutesOfMeetingTablesAsync(
        string connectionString,
        string schemaName,
        string tableNamePrefix,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();
        var allTableResults = new List<(string tableName, int inserted, int updated, int failed, List<string> errors)>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        try
        {
            // Step 1: Find matching tables
            List<string> matchingTables;
            
            // If the user explicitly selects a specific MOM table, only migrate to that table.
            if (string.Equals(tableNamePrefix, "MOM", StringComparison.OrdinalIgnoreCase) || 
                string.Equals(tableNamePrefix, "MinutesOfMeeting", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableNamePrefix, "MOM_Minutes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(tableNamePrefix, "MOM_Attendees", StringComparison.OrdinalIgnoreCase))
            {
                matchingTables = new List<string> { tableNamePrefix };
            }
            else
            {
                var momTables = await GetTablesWithPrefixAsync(connection, schemaName, "MinutesOfMeeting", cancellationToken);
                var momShortTables = await GetTablesWithPrefixAsync(connection, schemaName, "MOM", cancellationToken);
                matchingTables = momTables.Union(momShortTables, StringComparer.OrdinalIgnoreCase).ToList();
            }

            if (matchingTables.Count == 0)
            {
                response.ErrorMessages.Add($"No tables found with prefix 'MinutesOfMeeting' or 'MOM' in schema '{schemaName}'.");
                return response;
            }

            // Step 2: Migrate to each matching table
            foreach (var targetTable in matchingTables)
            {
                DataTable tableSpecificData = excelData;

                // Special row filtering for MOM child tables based on uuu_tab_id
                if (string.Equals(targetTable, "MOM_Attendees", StringComparison.OrdinalIgnoreCase))
                {
                    if (excelData.Columns.Contains("uuu_tab_id"))
                    {
                        var filteredRows = excelData.AsEnumerable()
                            .Where(r => r["uuu_tab_id"]?.ToString() == "1")
                            .ToList();

                        if (filteredRows.Count > 0)
                            tableSpecificData = filteredRows.CopyToDataTable();
                        else
                            continue; // Skip table if no matching rows
                    }
                }
                else if (string.Equals(targetTable, "MOM_Minutes", StringComparison.OrdinalIgnoreCase))
                {
                    if (excelData.Columns.Contains("uuu_tab_id"))
                    {
                        var filteredRows = excelData.AsEnumerable()
                            .Where(r => r["uuu_tab_id"]?.ToString() == "0")
                            .ToList();

                        if (filteredRows.Count > 0)
                            tableSpecificData = filteredRows.CopyToDataTable();
                        else
                            continue; // Skip table if no matching rows
                    }
                    else continue;
                }

                var tableResponse = await MigrateToSingleTableAsync(
                    connection,
                    schemaName,
                    targetTable,
                    tableSpecificData,
                    null,
                    cancellationToken);

                allTableResults.Add((
                    targetTable,
                    tableResponse.RowsInserted,
                    tableResponse.RowsUpdated,
                    tableResponse.RowsFailed,
                    tableResponse.ErrorMessages.ToList()
                ));

                response.RowErrors.AddRange(tableResponse.RowErrors);
            }

            // Step 3: Aggregate results
            response.Success = allTableResults.All(r => r.errors.Count == 0) && response.RowErrors.Count == 0;
            if (allTableResults.Count > 0)
            {
                response.RowsInserted = allTableResults.Sum(r => r.inserted);
                response.RowsUpdated = allTableResults.Sum(r => r.updated);
                response.RowsFailed = allTableResults.Sum(r => r.failed) + response.RowErrors.Count;

                var totalProcessed = response.RowsInserted + response.RowsUpdated;
                if (totalProcessed > 0)
                {
                    response.Message = $"Migrated to {allTableResults.Count} table(s). Total: {totalProcessed} row(s) processed.";
                }
            }

            foreach (var result in allTableResults.Where(r => r.errors.Count > 0))
            {
                response.ErrorMessages.Add($"Table '{result.tableName}': {string.Join("; ", result.errors)}");
            }
        }
        catch (Exception ex)
        {
            response.ErrorMessages.Add($"Error during MinutesOfMeeting migration: {ex.Message}");
        }

        return response;
    }


    private async Task<List<string>> GetTablesWithPrefixAsync(
        SqlConnection connection,
        string schemaName,
        string tablePrefix,
        CancellationToken cancellationToken)
    {
        var tables = new List<string>();

        var query = @"
            SELECT TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_SCHEMA = @SchemaName
                AND TABLE_TYPE = 'BASE TABLE'
                AND TABLE_NAME LIKE @TablePrefix + '%'
            ORDER BY TABLE_NAME";

        await using var command = new SqlCommand(query, connection);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@SchemaName", schemaName);
        command.Parameters.AddWithValue("@TablePrefix", tablePrefix);

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            tables.Add(reader.GetString(0));
        }

        return tables;
    }

    private async Task<UploadResponse> MigrateToSingleTableAsync(
        SqlConnection connection,
        string schemaName,
        string tableName,
        DataTable excelData,
        string? attachmentRecordType = null,
        CancellationToken cancellationToken = default,
        string? mappingTableName = null)
    {
        var response = new UploadResponse();
        var isOtBankGuaranteeMigration =
            string.Equals(mappingTableName, "OTBankGuarantee", StringComparison.OrdinalIgnoreCase);

        // Use a separate transaction for each table to ensure isolation
        var transaction = connection.BeginTransaction();
        var tempTableName = $"#TMP_{Guid.NewGuid():N}";

        try
        {
            // Step 1: Get target table metadata
            var tableMetadata = await GetTableMetadataAsync(connection, transaction, schemaName, tableName, cancellationToken);

            if (tableMetadata.Count == 0)
            {
                response.ErrorMessages.Add($"Table '{schemaName}.{tableName}' not found or has no columns.");
                transaction.Rollback();
                return response;
            }

            // Step 2: Match Excel columns to SQL columns
            var columnMappings = MatchColumns(excelData, tableMetadata, mappingTableName ?? tableName, attachmentRecordType);

            // ElectricalInstrumentationDBO child tables require parent FK ElectricalInstrumentationDBOID.
            // Ensure it is mapped from a parent id-like Excel column when dynamic matching doesn't add it.
            if (tableName.StartsWith("ElectricalInstrumentationDBO_", StringComparison.OrdinalIgnoreCase))
            {
                var electricalParentFk = tableMetadata.FirstOrDefault(m =>
                    string.Equals(m.ColumnName, "ElectricalInstrumentationDBOID", StringComparison.OrdinalIgnoreCase));

                var isElectricalParentFkMapped = columnMappings.Any(m =>
                    string.Equals(m.SqlColumnName, "ElectricalInstrumentationDBOID", StringComparison.OrdinalIgnoreCase));

                if (electricalParentFk != null && !isElectricalParentFkMapped)
                {
                    var parentIdExcelColumn = excelData.Columns.Cast<DataColumn>().FirstOrDefault(c =>
                        string.Equals(c.ColumnName, "id", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(c.ColumnName, "record_id", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(c.ColumnName, "ElectricalInstrumentationDBOID", StringComparison.OrdinalIgnoreCase));

                    if (parentIdExcelColumn != null)
                    {
                        columnMappings.Insert(0, new ColumnMapping
                        {
                            ExcelColumnName = parentIdExcelColumn.ColumnName,
                            SqlColumnName = electricalParentFk.ColumnName,
                            SqlDataType = electricalParentFk.DataType,
                            IsIdentity = electricalParentFk.IsIdentity,
                            IsNullable = electricalParentFk.IsNullable,
                            ForeignKeyTableSchema = electricalParentFk.ForeignKeyTableSchema,
                            ForeignKeyTableName = electricalParentFk.ForeignKeyTableName,
                            ForeignKeyColumnName = electricalParentFk.ForeignKeyColumnName,
                            ForeignKeyLookupColumnName = electricalParentFk.ForeignKeyLookupColumnName
                        });
                    }
                }
            }

            // SPECIAL RULE: Ensure the first Excel column is mapped to the Primary Key
            // This supports the requirement: "use the value from the first column of the Excel sheet as the primary key"
            // EXCEPTION: Skip this rule for BPComments and BPAttachments where IDs should typically be auto-generated to avoid conflicts (e.g. -1 values)
            bool shouldAutoMapPk = !string.Equals(tableName, "BPComments", StringComparison.OrdinalIgnoreCase) && 
                                   !string.Equals(tableName, "BPAttachments", StringComparison.OrdinalIgnoreCase);

            if (shouldAutoMapPk && excelData.Columns.Count > 0)
            {
                var primaryKeyColumn = tableMetadata.FirstOrDefault(m => m.IsPrimaryKey);
                if (primaryKeyColumn != null)
                {
                    // Check if PK is already mapped
                    var isPkMapped = columnMappings.Any(m => m.SqlColumnName.Equals(primaryKeyColumn.ColumnName, StringComparison.OrdinalIgnoreCase));
                    
                    if (!isPkMapped)
                    {
                        // Map first Excel column to Primary Key
                        columnMappings.Insert(0, new ColumnMapping
                        {
                            ExcelColumnName = excelData.Columns[0].ColumnName,
                            SqlColumnName = primaryKeyColumn.ColumnName,
                            SqlDataType = primaryKeyColumn.DataType,
                            IsIdentity = primaryKeyColumn.IsIdentity,
                            IsNullable = primaryKeyColumn.IsNullable
                        });
                    }
                }
            }

            if (columnMappings.Count == 0)
            {
                response.ErrorMessages.Add($"No matching columns found between Excel and SQL table '{tableName}'.");
                transaction.Rollback();
                return response;
            }

            // Step 3: Check for identity column
            var identityColumn = tableMetadata.FirstOrDefault(m => m.IsIdentity);
            var hasIdentityInExcel = identityColumn != null &&
                                    columnMappings.Any(m => m.SqlColumnName.Equals(identityColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

            // Step 4: Create temp table
            await CreateTempTableAsync(connection, transaction, tempTableName, tableMetadata, cancellationToken);

            // Step 5: Prepare DataTable with only matched columns
            var (mappedDataTable, rowErrors) = await PrepareMappedDataTableAsync(connection, transaction, excelData, columnMappings, tableName, schemaName, tableMetadata, cancellationToken);

            // OTBankGuarantee-specific behavior:
            // - Do not upsert
            // - Always generate sequential PK values from current table MAX(PK)
            if (isOtBankGuaranteeMigration && mappedDataTable.Rows.Count > 0)
            {
                var primaryKeyColumn = tableMetadata.FirstOrDefault(m => m.IsPrimaryKey);
                if (primaryKeyColumn == null)
                {
                    response.ErrorMessages.Add($"Primary key column not found for table '{schemaName}.{tableName}'.");
                    transaction.Rollback();
                    return response;
                }

                await AssignSequentialPrimaryKeysAsync(
                    connection,
                    transaction,
                    schemaName,
                    tableName,
                    primaryKeyColumn.ColumnName,
                    mappedDataTable,
                    cancellationToken);
            }

            // Add row errors to response
            response.RowErrors.AddRange(rowErrors);

            // Step 6: Bulk copy to temp table
            var rowsCopiedToTemp = await BulkCopyToTempTableAsync(
                connection,
                transaction,
                tempTableName,
                mappedDataTable,
                columnMappings,
                hasIdentityInExcel,
                cancellationToken);

            // Step 7: Get primary key columns
            var primaryKeyColumns = tableMetadata.Where(m => m.IsPrimaryKey).Select(m => m.ColumnName).ToList();

            int rowsInserted;
            int rowsUpdated;

            // Step 8:
            // - OTBankGuarantee: insert-only
            // - Other tables: regular upsert (MERGE)
            if (isOtBankGuaranteeMigration)
            {
                rowsInserted = await InsertFromTempToTargetAsync(
                    connection,
                    transaction,
                    schemaName,
                    tableName,
                    tempTableName,
                    columnMappings,
                    identityColumn,
                    hasIdentityInExcel,
                    cancellationToken);
                rowsUpdated = 0;
            }
            else
            {
                (rowsInserted, rowsUpdated) = await MergeFromTempToTargetAsync(
                    connection,
                    transaction,
                    schemaName,
                    tableName,
                    tempTableName,
                    columnMappings,
                    primaryKeyColumns,
                    identityColumn,
                    hasIdentityInExcel,
                    cancellationToken);
            }

            transaction.Commit();

            response.Success = rowErrors.Count == 0;
            response.RowsInserted = rowsInserted;
            response.RowsUpdated = rowsUpdated;
            response.RowsFailed = rowErrors.Count;

        }
        catch (Exception ex)
        {
            try
            {
                transaction.Rollback();
            }
            catch
            {
                // Ignore rollback errors if transaction already completed
            }
            
            response.ErrorMessages.Add($"Error migrating to table '{tableName}': {ex.Message}");
            if (ex.InnerException != null)
            {
                response.ErrorMessages.Add($"Inner exception: {ex.InnerException.Message}");
            }
        }
        finally
        {
            // Clean up temp table - pass null for transaction as it's already committed or rolled back
            try
            {
                await DropTempTableAsync(connection, null, tempTableName, cancellationToken);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        return response;
    }

    private async Task AssignSequentialPrimaryKeysAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string primaryKeyColumnName,
        DataTable mappedDataTable,
        CancellationToken cancellationToken)
    {
        if (!mappedDataTable.Columns.Contains(primaryKeyColumnName))
        {
            throw new InvalidOperationException($"Mapped data does not contain primary key column '{primaryKeyColumnName}'.");
        }

        var getMaxIdQuery = $@"
            SELECT ISNULL(MAX([{primaryKeyColumnName}]), 0)
            FROM [{schemaName}].[{tableName}] WITH (UPDLOCK, HOLDLOCK)";

        await using var command = new SqlCommand(getMaxIdQuery, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        var currentMaxObj = await command.ExecuteScalarAsync(cancellationToken);
        var currentMax = currentMaxObj == null || currentMaxObj == DBNull.Value ? 0L : Convert.ToInt64(currentMaxObj);

        foreach (DataRow row in mappedDataTable.Rows)
        {
            currentMax++;
            row[primaryKeyColumnName] = currentMax;
        }
    }

    private async Task<List<ColumnMetadata>> GetTableMetadataAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        CancellationToken cancellationToken)
    {
        var metadata = new List<ColumnMetadata>();

        var query = @"
            SELECT 
                c.COLUMN_NAME,
                c.DATA_TYPE,
                c.IS_NULLABLE,
                c.CHARACTER_MAXIMUM_LENGTH,
                c.NUMERIC_PRECISION,
                c.NUMERIC_SCALE,
                CASE WHEN ic.OBJECT_ID IS NOT NULL THEN 1 ELSE 0 END AS IS_IDENTITY,
                CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PRIMARY_KEY
            FROM INFORMATION_SCHEMA.COLUMNS c
            LEFT JOIN sys.identity_columns ic ON ic.object_id = OBJECT_ID(@SchemaTable) 
                AND ic.name = c.COLUMN_NAME
            LEFT JOIN (
                SELECT ku.TABLE_SCHEMA, ku.TABLE_NAME, ku.COLUMN_NAME
                FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                    ON tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
                    AND tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                    AND tc.TABLE_SCHEMA = ku.TABLE_SCHEMA
                    AND tc.TABLE_NAME = ku.TABLE_NAME
            ) pk ON pk.TABLE_SCHEMA = c.TABLE_SCHEMA 
                AND pk.TABLE_NAME = c.TABLE_NAME 
                AND pk.COLUMN_NAME = c.COLUMN_NAME
            WHERE c.TABLE_SCHEMA = @SchemaName 
                AND c.TABLE_NAME = @TableName
            ORDER BY c.ORDINAL_POSITION";

        await using var command = new SqlCommand(query, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@SchemaName", schemaName);
        command.Parameters.AddWithValue("@TableName", tableName);
        command.Parameters.AddWithValue("@SchemaTable", $"{schemaName}.{tableName}");

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            metadata.Add(new ColumnMetadata
            {
                ColumnName = reader.GetString(0),
                DataType = reader.GetString(1),
                IsNullable = reader.GetString(2) == "YES",
                MaxLength = reader.IsDBNull(3) ? (int?)null : Convert.ToInt32(reader.GetValue(3)),
                NumericPrecision = reader.IsDBNull(4) ? (int?)null : Convert.ToInt32(reader.GetValue(4)),
                NumericScale = reader.IsDBNull(5) ? (int?)null : Convert.ToInt32(reader.GetValue(5)),
                IsIdentity = Convert.ToInt32(reader.GetValue(6)) == 1,
                IsPrimaryKey = Convert.ToInt32(reader.GetValue(7)) == 1
            });
        }

        return metadata;
    }

    private async Task<string?> FindLookupColumnAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string? parentTableSchema,
        string? parentTableName,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(parentTableSchema) || string.IsNullOrWhiteSpace(parentTableName))
            return null;

        // Common column names to search for (in order of preference)
        var lookupColumnNames = new[] { "Name", "ContactName", "Description", "Title", "DisplayName", "FullName" };

        var query = @"
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = @SchemaName
                AND TABLE_NAME = @TableName
                AND COLUMN_NAME IN ('Name', 'ContactName', 'Description', 'Title', 'DisplayName', 'FullName')
            ORDER BY CASE COLUMN_NAME
                WHEN 'Name' THEN 1
                WHEN 'ContactName' THEN 2
                WHEN 'Description' THEN 3
                WHEN 'Title' THEN 4
                WHEN 'DisplayName' THEN 5
                WHEN 'FullName' THEN 6
                ELSE 99
            END";

        try
        {
            await using var command = new SqlCommand(query, connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            command.Parameters.AddWithValue("@SchemaName", parentTableSchema);
            command.Parameters.AddWithValue("@TableName", parentTableName);

            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                return reader.GetString(0);
            }
        }
        catch
        {
            // If lookup fails, return null - we'll try direct ID conversion
        }

        return null;
    }

    private bool IsZeroValue(object value)
    {
        if (value == null || value == DBNull.Value)
            return false;

        return value switch
        {
            int intVal => intVal == 0,
            long longVal => longVal == 0,
            short shortVal => shortVal == 0,
            byte byteVal => byteVal == 0,
            decimal decimalVal => decimalVal == 0,
            double doubleVal => doubleVal == 0,
            float floatVal => floatVal == 0,
            string strVal => strVal == "0" || string.IsNullOrWhiteSpace(strVal),
            _ => false
        };
    }

    private async Task<bool> ValidateForeignKeyValueAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string fkTableSchema,
        string fkTableName,
        string fkColumnName,
        object fkValue,
        CancellationToken cancellationToken)
    {
        try
        {
            var query = $@"
                SELECT COUNT(1)
                FROM [{fkTableSchema}].[{fkTableName}]
                WHERE [{fkColumnName}] = @FkValue";

            await using var command = new SqlCommand(query, connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            command.Parameters.AddWithValue("@FkValue", fkValue);

            var count = await command.ExecuteScalarAsync(cancellationToken);
            return count != null && Convert.ToInt32(count) > 0;
        }
        catch
        {
            // If validation fails (e.g., table doesn't exist), return false to skip the row
            return false;
        }
    }

    private async Task<object?> LookupForeignKeyValueAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string parentTableSchema,
        string parentTableName,
        string parentKeyColumn,
        string? lookupColumnName,
        object excelValue,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return DBNull.Value;

        var lookupValue = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(lookupValue))
            return DBNull.Value;

        try
        {
            // First, try to see if the Excel value is already a numeric ID
            // If it's numeric, try direct lookup by ID first
            object? numericIdValue = null;
            if (long.TryParse(lookupValue, out var numericId))
            {
                numericIdValue = numericId;
            }
            else if (int.TryParse(lookupValue, out var intId))
            {
                numericIdValue = intId;
            }

            if (numericIdValue != null)
            {
                var directQuery = $@"
                    SELECT TOP 1 [{parentKeyColumn}]
                    FROM [{parentTableSchema}].[{parentTableName}]
                    WHERE [{parentKeyColumn}] = @DirectId";

                await using var directCommand = new SqlCommand(directQuery, connection, transaction);
                directCommand.CommandTimeout = SqlCommandTimeout;
                directCommand.Parameters.AddWithValue("@DirectId", numericIdValue);

                var directResult = await directCommand.ExecuteScalarAsync(cancellationToken);
                if (directResult != null && directResult != DBNull.Value)
                {
                    return directResult;
                }
            }

            // If direct ID lookup failed or value is not numeric, try lookup by name/description
            if (!string.IsNullOrWhiteSpace(lookupColumnName))
            {
                var nameQuery = $@"
                    SELECT TOP 1 [{parentKeyColumn}]
                    FROM [{parentTableSchema}].[{parentTableName}]
                    WHERE [{lookupColumnName}] = @LookupValue
                    ORDER BY [{parentKeyColumn}]";

                await using var nameCommand = new SqlCommand(nameQuery, connection, transaction);
                nameCommand.CommandTimeout = SqlCommandTimeout;
                nameCommand.Parameters.AddWithValue("@LookupValue", lookupValue);

                var nameResult = await nameCommand.ExecuteScalarAsync(cancellationToken);
                if (nameResult != null && nameResult != DBNull.Value)
                {
                    return nameResult;
                }
            }

            // If both lookups failed, return DBNull (will be handled as conversion error)
            return DBNull.Value;
        }
        catch
        {
            // If lookup fails, return DBNull - let the conversion handle it
            // (might be a direct ID that needs type conversion)
            return DBNull.Value;
        }
    }

    private DataTable FilterExcelDataByTabId(DataTable excelData, string columnName, string filterValue)
    {
        // Create a new DataTable with the same structure
        var filteredData = excelData.Clone();

        // Check if the filter column exists
        if (!excelData.Columns.Contains(columnName))
        {
            // If column doesn't exist, return all data (no filtering)
            foreach (DataRow row in excelData.Rows)
            {
                filteredData.ImportRow(row);
            }
            return filteredData;
        }

        // Filter rows where the column value matches the filter value
        foreach (DataRow row in excelData.Rows)
        {
            var cellValue = row[columnName]?.ToString()?.Trim() ?? string.Empty;
            if (string.Equals(cellValue, filterValue, StringComparison.OrdinalIgnoreCase))
            {
                filteredData.ImportRow(row);
            }
        }

        return filteredData;
    }

    private async Task<UploadResponse> MigrateOrderTransmittalLineItemBankGuaranteeAsync(
        string connectionString,
        string schemaName,
        DataTable excelData,
        string? attachmentRecordType = null,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        if (!excelData.Columns.Contains("uuu_tab_id"))
        {
            response.ErrorMessages.Add("Column 'uuu_tab_id' is required for OrderTransmittalLineItemBankGuarantee migration.");
            return response;
        }

        var filteredData = FilterExcelDataByTabId(excelData, "uuu_tab_id", "6");
        if (filteredData.Rows.Count == 0)
        {
            response.ErrorMessages.Add("No rows found with uuu_tab_id = 6 for OrderTransmittalLineItemBankGuarantee migration.");
            return response;
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        return await MigrateToSingleTableAsync(
            connection,
            schemaName,
            "BankGuarantee",
            filteredData,
            attachmentRecordType,
            cancellationToken,
            "OTBankGuarantee");
    }

    private List<ColumnMapping> MatchColumns(DataTable excelData, List<ColumnMetadata> tableMetadata, string tableName, string? attachmentRecordType = null)
    {
        var mappings = new List<ColumnMapping>();

        // Check if this is CommunicationProtocol table - use hardcoded mapping
        if (string.Equals(tableName, "CommunicationProtocol", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForCommunicationProtocol(excelData, tableMetadata);
        }

        // Check if this is BankGuarantee table - use hardcoded mapping
        if (string.Equals(tableName, "BankGuarantee", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForBankGuarantee(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "OTBankGuarantee", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForOrderTransmittalLineItemBankGuarantee(excelData, tableMetadata);
        }

        // Check if this is BankGuaranteeSpares table - use hardcoded mapping
        if (string.Equals(tableName, "BankGuaranteeSpares", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForBankGuaranteeSpares(excelData, tableMetadata);
        }

        // Check if this is DeviationRequest table - use hardcoded mapping
        if (string.Equals(tableName, "DeviationRequest", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForDeviationRequest(excelData, tableMetadata);
        }

        // Check if this is DesignRecommendation table - use hardcoded mapping
        if (string.Equals(tableName, "DesignRecommendation", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForDesignRecommendation(excelData, tableMetadata);
        }

        // Check if this is FrameMaster table - use hardcoded mapping
        if (string.Equals(tableName, "FrameMaster", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForFrameMaster(excelData, tableMetadata);
        }

        // Check if this is BankMaster table - use hardcoded mapping
        if (string.Equals(tableName, "BankMaster", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForBankMaster(excelData, tableMetadata);
        }

        // Check if this is UserList table - use hardcoded mapping
        if (string.Equals(tableName, "UserList", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForUserList(excelData, tableMetadata);
        }

        // Check if this is MonthlyActualCollection table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyActualCollection", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyActualCollection(excelData, tableMetadata);
        }

        // Check if this is InitialCashFlowPlan table - use hardcoded mapping
        if (string.Equals(tableName, "InitialCashFlowPlan", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("InitialCashFlowPlan", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForInitialCashFlowPlan(excelData, tableMetadata);
        }

        // Check if this is MonthlyPlanning table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyPlanning", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyPlanning(excelData, tableMetadata);
        }

        // Check if this is MonthlyPlanningLineItem table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyPlanningLineItem", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tableName, "MonthlyPlanning_LineItem", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("MonthlyPlanningLineItem", StringComparison.OrdinalIgnoreCase) ||
            tableName.StartsWith("MonthlyPlanning_LineItem", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyPlanningLineItem(excelData, tableMetadata);
        }

        // Check if this is MonthlyActualCollectionPlanned table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyActualCollectionPlanned", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyActualCollectionPlanned(excelData, tableMetadata);
        }

        // Check if this is MonthlyActualUnplannedCollection table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyActualUnplannedCollection", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyActualUnplannedCollection(excelData, tableMetadata);
        }

        // Check if this is SparesOrderTransmittal table - use hardcoded mapping
        if (string.Equals(tableName, "SparesOrderTransmittal", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForSparesOrderTransmittal(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "SparesOrderTransmittalLineItem", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForSparesOrderTransmittalLineItem(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "OrderReceiptAcknowledgement", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForOrderReceiptAcknowledgement(excelData, tableMetadata);
        }



        if (string.Equals(tableName, "AuditAction", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForAuditAction(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "AuditLog", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForAuditLog(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "RCCA", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForRCCA(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "RCCA_StandardLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForRCCA_StandardLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "RCCA_SelectTeamMembersLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForRCCA_SelectTeamMembersLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_MajorMilestoneLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_MajorMilestoneLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_ScopeOfSupplyLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_ScopeOfSupplyLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_ProcurementProgressofBoughtOutsItemLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_ProcurementProgressofBoughtOutsItemLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_TurbineManufacturingProgressLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_TurbineManufacturingProgressLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_InspectionDispatchPlanLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_InspectionDispatchPlanLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_CashInFlowPlanLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_CashInFlowPlanLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_LookAheadTaskforNext30DaysLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_LookAheadTaskforNext30DaysLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_EngineeringProgressLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_EngineeringProgressLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_InputsRequiredFromCustomerLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_InputsRequiredFromCustomerLI(excelData, tableMetadata);
        }

        if (string.Equals(tableName, "MonthlyProgressReport_DBOSummarizeSheetLI", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyProgressReport_DBOSummarizeSheetLI(excelData, tableMetadata);
        }

        // Check if this is SpecificationRelease table - use hardcoded mapping
        if (string.Equals(tableName, "SpecificationRelease", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForSpecificationRelease(excelData, tableMetadata);
        }


        // Check if this is ContractClearance table - use hardcoded mapping
        if (string.Equals(tableName, "ContractClearance", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForContractClearance(excelData, tableMetadata);
        }

        // Check if this is AdditionalOrderBooking table - use hardcoded mapping
        if (string.Equals(tableName, "AdditionalOrderBooking", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForAdditionalOrderBooking(excelData, tableMetadata);
        }

        // Check if this is ContractOnHold table - use hardcoded mapping
        if (tableName.StartsWith("ContractOnHold", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForContractOnHold(excelData, tableMetadata);
        }

        // Check if this is LCReview table - use hardcoded mapping
        if (string.Equals(tableName, "LCReview", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForLCReview(excelData, tableMetadata);
        }

        // Check if this is LCReview_NotesObservation table - use hardcoded mapping
        if (string.Equals(tableName, "LCReview_NotesObservation", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForLCReviewNotesObservation(excelData, tableMetadata);
        }

        // Check if this is InitialCashPlan table - use hardcoded mapping
        if (string.Equals(tableName, "InitialCashPlan", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForInitialCashPlan(excelData, tableMetadata);
        }

        // Check if this is payment_supply table - use hardcoded mapping
        if (string.Equals(tableName, "payment_supply", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForPaymentSupply(excelData, tableMetadata);
        }

        // Check if this is payment_enc table - use hardcoded mapping
        if (string.Equals(tableName, "payment_enc", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForPaymentENC(excelData, tableMetadata);
        }

        // Check if this is LiquidatedDamage table - use hardcoded mapping
        if (string.Equals(tableName, "LiquidatedDamage", StringComparison.OrdinalIgnoreCase) || string.Equals(tableName, "liquidated_damage", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForLiquidatedDamage(excelData, tableMetadata);
        }

        // Check if table name starts with "MinutesOfMeeting" or "MOM" - use hardcoded mapping
        if (tableName.StartsWith("MinutesOfMeeting", StringComparison.OrdinalIgnoreCase) || 
            tableName.StartsWith("MOM", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(tableName, "MOM_Attendees", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForMOMAttendees(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "MOM_Minutes", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForMOMMinutes(excelData, tableMetadata);
            }
            return MatchColumnsForMinutesOfMeeting(excelData, tableMetadata);
        }

        // Check if this is LetterOfCorrespondence table - use hardcoded mapping
        if (string.Equals(tableName, "LetterOfCorrespondence", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForLetterOfCorrespondence(excelData, tableMetadata);
        }

        // Check if this is CustomerMaster table - use hardcoded mapping
        if (string.Equals(tableName, "CustomerMaster", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForCustomerMaster(excelData, tableMetadata);
        }

        // Check if this is CustomerContacts table - use hardcoded mapping
        if (string.Equals(tableName, "CustomerContacts", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForCustomerContacts(excelData, tableMetadata);
        }

        // Check if this is VendorMaster table - use hardcoded mapping
        if (string.Equals(tableName, "VendorMaster", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForVendorMaster(excelData, tableMetadata);
        }

        // Check if this is BPAttachments table - use hardcoded mapping with dynamic selection based on AttachmentRecordType
        if (string.Equals(tableName, "BPAttachments", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForBPAttachments(excelData, tableMetadata, attachmentRecordType);
        }

        // Check if this is Project table - use hardcoded mapping
        if (string.Equals(tableName, "Project", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForProject(excelData, tableMetadata);
        }

        // Check if this is Risks table - use hardcoded mapping
        if (string.Equals(tableName, "Risks", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForRisks(excelData, tableMetadata);
        }

        // Check if this is Risk_MitigationPlan table - use hardcoded mapping
        if (string.Equals(tableName, "Risk_MitigationPlan", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForRiskMitigationPlan(excelData, tableMetadata);
        }

        // Check if this is PositiveRecall table - use hardcoded mapping
        if (string.Equals(tableName, "PositiveRecall", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForPositiveRecall(excelData, tableMetadata);
        }

        // Check if this is MonthlyCashInflow table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyCashInflow", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyCashInflow(excelData, tableMetadata);
        }

        // Check if this is MonthlyCashInflow_CustomerCare table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyCashInflow_CustomerCare", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyCashInflowCustomerCare(excelData, tableMetadata);
        }

        // Check if this is MonthlyCashInflow_MarketingAdvance table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyCashInflow_MarketingAdvance", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyCashInflowMarketingAdvance(excelData, tableMetadata);
        }

        // Check if this is MonthlyCashInflow_ICDue table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyCashInflow_ICDue", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyCashInflowICDue(excelData, tableMetadata);
        }

        // Check if this is MonthlyCashInflow_ICNotDue table - use hardcoded mapping
        if (string.Equals(tableName, "MonthlyCashInflow_ICNotDue", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMonthlyCashInflowICNotDue(excelData, tableMetadata);
        }

        // Check if this is HandoverProtocol table - use hardcoded mapping
        if (string.Equals(tableName, "HandoverProtocol", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForHandoverProtocol(excelData, tableMetadata);
        }

        // Check if this is HandoverProtocolLineItem table - use hardcoded mapping
        if (string.Equals(tableName, "HandoverProtocolLineItem", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForHandoverProtocolLineItem(excelData, tableMetadata);
        }

        // Check if this is InternalAudit table - use hardcoded mapping
        if (string.Equals(tableName, "InternalAudit", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForInternalAudit(excelData, tableMetadata);
        }

        // Check if this is InternalAudit_Notes table - use hardcoded mapping
        if (string.Equals(tableName, "InternalAudit_Notes", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForInternalAuditNotes(excelData, tableMetadata);
        }

        // Check if this is InternalAudit_ObservationDetails table - use hardcoded mapping
        if (string.Equals(tableName, "InternalAudit_ObservationDetails", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForInternalAuditObservationDetails(excelData, tableMetadata);
        }

        // Check if this is OTRisk table - use hardcoded mapping
        if (string.Equals(tableName, "OTRisk", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForOTRisk(excelData, tableMetadata);
        }

        // Check if this is BPComments table - use hardcoded mapping
        if (string.Equals(tableName, "BPComments", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForBPComments(excelData, tableMetadata, attachmentRecordType);
        }

        // Check if table name starts with "Turbine" - use hardcoded mapping
        if (tableName.StartsWith("Turbine", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForTurbine(excelData, tableMetadata);
        }

        // Check if table name starts with "MechanicalDBO" - use hardcoded mapping
        if (tableName.StartsWith("MechanicalDBO", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForMechanicalDBO(excelData, tableMetadata);
        }

        // Check if table name starts with "OrderTransmittal" - use hardcoded mapping
        if (tableName.StartsWith("OrderTransmittal", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(tableName, "OrderTransmittal_Notes", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalNotes(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_ContractualDeliveryMapping", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalContractualDeliveryMapping(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_DefectLiability", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalDefectLiability(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_FormatsReviewed", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalFormatsReviewed(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_GoverningLaw", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalGoverningLaw(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_ScopeOfWork", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalScopeOfWork(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_RelevantTerms", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalRelevantTerms(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_PerformanceGuarantee", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalPerformanceGuarantee(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_LiquidatedDamages", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalLiquidatedDamages(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_PaymentDetails", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalPaymentDetails(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_TermsOfDelivery", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalTermsOfDelivery(excelData, tableMetadata);
            }
            if (string.Equals(tableName, "OrderTransmittal_CostSheetReview", StringComparison.OrdinalIgnoreCase))
            {
                return MatchColumnsForOrderTransmittalCostSheetReview(excelData, tableMetadata);
            }
            return MatchColumnsForOrderTransmittal(excelData, tableMetadata);
        }

        // Use hardcoded mapping only for ElectricalInstrumentationDBO parent table.
        // Child tables should use dynamic matching against their own columns.
        if (string.Equals(tableName, "ElectricalInstrumentationDBO", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForElectricalInstrumentationDBO(excelData, tableMetadata);
        }
        if (tableName.StartsWith("ApprovalLog", StringComparison.OrdinalIgnoreCase))
        {
            return MatchColumnsForApprovalLog(excelData, tableMetadata);
        }
        // For other tables, use existing dynamic matching logic
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        foreach (var sqlColumn in tableMetadata)
        {
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(sqlColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn != null)
            {
                mappings.Add(new ColumnMapping
                {
                    ExcelColumnName = excelColumn.ColumnName,
                    SqlColumnName = sqlColumn.ColumnName,
                    SqlDataType = sqlColumn.DataType,
                    IsIdentity = sqlColumn.IsIdentity,
                    IsNullable = sqlColumn.IsNullable,
                    ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                    ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                    ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                    ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
                });
            }
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForCommunicationProtocol(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in CommunicationProtocolColumnMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }



    private List<ColumnMapping> MatchColumnsForMinutesOfMeeting(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in MinutesOfMeetingMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }


    private List<ColumnMapping> MatchColumnsForMOMAttendees(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        foreach (var mappingEntry in MOMAttendeesMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMOMMinutes(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        foreach (var mappingEntry in MOMMinutesMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForContractClearance(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in ContractClearanceMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForAdditionalOrderBooking(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in AdditionalOrderBookingMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForContractOnHold(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in ContractOnHoldMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForLCReview(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in LCReviewMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForLCReviewNotesObservation(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in LCReviewNotesObservationMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForInitialCashPlan(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in InitialCashPlanMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }






    private List<ColumnMapping> MatchColumnsForPaymentSupply(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in PaymentSupplyMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForLiquidatedDamage(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in LiquidatedDamageMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForPaymentENC(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in PaymentENCMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForLetterOfCorrespondence(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in LetterOfCorrespondenceMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForBankGuarantee(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in BankGuaranteeMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalLineItemBankGuarantee(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        void AddMappingsFromDictionary(Dictionary<string, string> sourceMapping)
        {
            foreach (var mappingEntry in sourceMapping)
            {
                var excelColumn = excelColumns.FirstOrDefault(
                    ec => ec.ColumnName.Equals(mappingEntry.Key, StringComparison.OrdinalIgnoreCase));

                if (excelColumn == null)
                    continue;

                if (!sqlColumnLookup.TryGetValue(mappingEntry.Value, out var sqlColumn))
                    continue;

                var alreadyMapped = mappings.Any(m =>
                    m.ExcelColumnName.Equals(excelColumn.ColumnName, StringComparison.OrdinalIgnoreCase) &&
                    m.SqlColumnName.Equals(sqlColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

                if (alreadyMapped)
                    continue;

                mappings.Add(new ColumnMapping
                {
                    ExcelColumnName = excelColumn.ColumnName,
                    SqlColumnName = sqlColumn.ColumnName,
                    SqlDataType = sqlColumn.DataType,
                    IsIdentity = sqlColumn.IsIdentity,
                    IsNullable = sqlColumn.IsNullable
                });
            }
        }

        // Use line-item specific mapping profile.
        AddMappingsFromDictionary(OrderTransmittalLineItemBankGuaranteeMapping);

        // Fallback for files that already use SQL-friendly headers.
        foreach (var sqlColumn in tableMetadata)
        {
            var isAlreadyMapped = mappings.Any(m =>
                m.SqlColumnName.Equals(sqlColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

            if (isAlreadyMapped)
                continue;

            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(sqlColumn.ColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForBankGuaranteeSpares(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, BankGuaranteeSparesMapping);
    }

    private List<ColumnMapping> MatchColumnsForDeviationRequest(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, DeviationRequestMapping);
    }

    private List<ColumnMapping> MatchColumnsForDesignRecommendation(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, DesignRecommendationMapping);
    }

    private List<ColumnMapping> MatchColumnsForFrameMaster(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, FrameMasterMapping);
    }

    private List<ColumnMapping> MatchColumnsForBankMaster(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, BankMasterMapping);
    }

    private List<ColumnMapping> MatchColumnsForUserList(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in UserListMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyActualCollection(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in MonthlyActualCollectionMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForInitialCashFlowPlan(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in InitialCashFlowPlanMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyPlanning(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in MonthlyPlanningMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyPlanningLineItem(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in MonthlyPlanningLineItemMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyActualCollectionPlanned(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();

        var excelColumnLookup = excelData.Columns.Cast<DataColumn>()
            .GroupBy(c => c.ColumnName.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var sqlColumnLookup = tableMetadata
            .ToDictionary(c => c.ColumnName, c => c, StringComparer.OrdinalIgnoreCase);

        var addedSqlColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in MonthlyActualCollectionPlannedMapping)
        {
            var excelColumnName = mapping.Key.Trim();
            var sqlColumnName   = mapping.Value;

            if (addedSqlColumns.Contains(sqlColumnName))
                continue;

            if (!excelColumnLookup.TryGetValue(excelColumnName, out var excelColumn))
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName   = sqlColumn.ColumnName,
                SqlDataType     = sqlColumn.DataType,
                IsIdentity      = sqlColumn.IsIdentity,
                IsNullable      = sqlColumn.IsNullable
            });

            addedSqlColumns.Add(sqlColumnName);
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyActualUnplannedCollection(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();

        var excelColumnLookup = excelData.Columns.Cast<DataColumn>()
            .GroupBy(c => c.ColumnName.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var sqlColumnLookup = tableMetadata
            .ToDictionary(c => c.ColumnName, c => c, StringComparer.OrdinalIgnoreCase);

        var addedSqlColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in MonthlyActualUnplannedCollectionMapping)
        {
            var excelColumnName = mapping.Key.Trim();
            var sqlColumnName   = mapping.Value;

            if (addedSqlColumns.Contains(sqlColumnName))
                continue;

            if (!excelColumnLookup.TryGetValue(excelColumnName, out var excelColumn))
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName   = sqlColumn.ColumnName,
                SqlDataType     = sqlColumn.DataType,
                IsIdentity      = sqlColumn.IsIdentity,
                IsNullable      = sqlColumn.IsNullable
            });

            addedSqlColumns.Add(sqlColumnName);
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForSpecificationRelease(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();

        var excelColumnLookup = excelData.Columns.Cast<DataColumn>()
            .GroupBy(c => c.ColumnName.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var sqlColumnLookup = tableMetadata
            .ToDictionary(c => c.ColumnName, c => c, StringComparer.OrdinalIgnoreCase);

        var addedSqlColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in SpecificationReleaseMapping)
        {
            var excelColumnName = mapping.Key.Trim();
            var sqlColumnName   = mapping.Value;

            if (addedSqlColumns.Contains(sqlColumnName))
                continue;

            if (!excelColumnLookup.TryGetValue(excelColumnName, out var excelColumn))
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName   = sqlColumn.ColumnName,
                SqlDataType     = sqlColumn.DataType,
                IsIdentity      = sqlColumn.IsIdentity,
                IsNullable      = sqlColumn.IsNullable,
                ForeignKeyTableSchema      = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName        = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName       = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });

            addedSqlColumns.Add(sqlColumnName);
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForSparesOrderTransmittal(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();

        var excelColumnLookup = excelData.Columns.Cast<DataColumn>()
            .GroupBy(c => c.ColumnName.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var sqlColumnLookup = tableMetadata
            .ToDictionary(c => c.ColumnName, c => c, StringComparer.OrdinalIgnoreCase);

        var addedSqlColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in SparesOrderTransmittalMapping)
        {
            var excelColumnName = mapping.Key.Trim();
            var sqlColumnName   = mapping.Value;

            if (addedSqlColumns.Contains(sqlColumnName))
                continue;

            if (!excelColumnLookup.TryGetValue(excelColumnName, out var excelColumn))
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName   = sqlColumn.ColumnName,
                SqlDataType     = sqlColumn.DataType,
                IsIdentity      = sqlColumn.IsIdentity,
                IsNullable      = sqlColumn.IsNullable,
                ForeignKeyTableSchema      = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName        = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName       = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });

            addedSqlColumns.Add(sqlColumnName);
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForSparesOrderTransmittalLineItem(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();

        var excelColumnLookup = excelData.Columns.Cast<DataColumn>()
            .GroupBy(c => c.ColumnName.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var sqlColumnLookup = tableMetadata
            .ToDictionary(c => c.ColumnName, c => c, StringComparer.OrdinalIgnoreCase);

        var addedSqlColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in SparesOrderTransmittalLineItemMapping)
        {
            var excelColumnName = mapping.Key.Trim();
            var sqlColumnName   = mapping.Value;

            if (addedSqlColumns.Contains(sqlColumnName))
                continue;

            if (!excelColumnLookup.TryGetValue(excelColumnName, out var excelColumn))
                continue;

            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue;

            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName   = sqlColumn.ColumnName,
                SqlDataType     = sqlColumn.DataType,
                IsIdentity      = sqlColumn.IsIdentity,
                IsNullable      = sqlColumn.IsNullable,
                ForeignKeyTableSchema      = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName        = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName       = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });

            addedSqlColumns.Add(sqlColumnName);
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForOrderReceiptAcknowledgement(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in OrderReceiptAcknowledgementMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
            {
                // FALLBACK: Try matching by SQL column name directly if mapped name is not found
                excelColumn = excelColumns.FirstOrDefault(
                    ec => ec.ColumnName.Equals(sqlColumnName, StringComparison.OrdinalIgnoreCase));
            }

            if (excelColumn == null)
                continue; // Skip if column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsUsingDictionary(DataTable excelData, List<ColumnMetadata> tableMetadata, Dictionary<string, string> mappingDict)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in mappingDict)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
            {
                // FALLBACK: Try matching by SQL column name directly if mapped name is not found
                excelColumn = excelColumns.FirstOrDefault(
                    ec => ec.ColumnName.Equals(sqlColumnName, StringComparison.OrdinalIgnoreCase));
            }

            if (excelColumn == null)
                continue; // Skip if column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalNotes(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalNotesMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalContractualDeliveryMapping(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalContractualDeliveryMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalDefectLiability(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalDefectLiabilityMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalFormatsReviewed(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalFormatsReviewedMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalGoverningLaw(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalGoverningLawMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalScopeOfWork(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalScopeOfWorkMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalRelevantTerms(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalRelevantTermsMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalPerformanceGuarantee(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalPerformanceGuaranteeMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalLiquidatedDamages(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalLiquidatedDamagesMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalPaymentDetails(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalPaymentDetailsMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalTermsOfDelivery(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalTermsOfDeliveryMapping);
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittalCostSheetReview(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OrderTransmittalCostSheetReviewMapping);
    }

    private List<ColumnMapping> MatchColumnsForRisks(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, RisksMapping);
    }

    private List<ColumnMapping> MatchColumnsForRiskMitigationPlan(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, RiskMitigationPlanMapping);
    }

    private List<ColumnMapping> MatchColumnsForPositiveRecall(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, PositiveRecallMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyCashInflow(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyCashInflowMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyCashInflowCustomerCare(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyCashInflowCustomerCareMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyCashInflowMarketingAdvance(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyCashInflowMarketingAdvanceMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyCashInflowICDue(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyCashInflowICDueMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyCashInflowICNotDue(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyCashInflowICNotDueMapping);
    }

    private List<ColumnMapping> MatchColumnsForHandoverProtocol(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, HandoverProtocolMapping);
    }

    private List<ColumnMapping> MatchColumnsForHandoverProtocolLineItem(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, HandoverProtocolLineItemMapping);
    }

    private List<ColumnMapping> MatchColumnsForInternalAudit(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, InternalAuditMapping);
    }

    private List<ColumnMapping> MatchColumnsForInternalAuditNotes(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, InternalAuditNotesMapping);
    }

    private List<ColumnMapping> MatchColumnsForInternalAuditObservationDetails(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, InternalAuditObservationDetailsMapping);
    }

    private List<ColumnMapping> MatchColumnsForOTRisk(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, OTRiskMapping);
    }

    private List<ColumnMapping> MatchColumnsForAuditAction(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, AuditActionMapping);
    }

    private List<ColumnMapping> MatchColumnsForAuditLog(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, AuditLogMapping);
    }

    private List<ColumnMapping> MatchColumnsForRCCA(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, RCCAMapping);
    }

    private List<ColumnMapping> MatchColumnsForRCCA_StandardLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, RCCA_StandardLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForRCCA_SelectTeamMembersLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, RCCA_SelectTeamMembersLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReportMapping);
        
        // Special case: mpr_title_sdt120 also maps to Title column if it exists in SQL
        var excelColumn = excelData.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "mpr_title_sdt120", StringComparison.OrdinalIgnoreCase));
        var sqlTitleColumn = tableMetadata.FirstOrDefault(m => string.Equals(m.ColumnName, "Title", StringComparison.OrdinalIgnoreCase));
        
        if (excelColumn != null && sqlTitleColumn != null && !mappings.Any(m => string.Equals(m.SqlColumnName, "Title", StringComparison.OrdinalIgnoreCase)))
        {
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlTitleColumn.ColumnName,
                SqlDataType = sqlTitleColumn.DataType,
                IsIdentity = sqlTitleColumn.IsIdentity,
                IsNullable = sqlTitleColumn.IsNullable
            });
        }
        
        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_MajorMilestoneLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_MajorMilestoneLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_ScopeOfSupplyLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_ScopeOfSupplyLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_ProcurementProgressofBoughtOutsItemLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_ProcurementProgressofBoughtOutsItemLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_TurbineManufacturingProgressLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_TurbineManufacturingProgressLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_InspectionDispatchPlanLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_InspectionDispatchPlanLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_CashInFlowPlanLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_CashInFlowPlanLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_LookAheadTaskforNext30DaysLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_LookAheadTaskforNext30DaysLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_EngineeringProgressLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_EngineeringProgressLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_InputsRequiredFromCustomerLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_InputsRequiredFromCustomerLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForMonthlyProgressReport_DBOSummarizeSheetLI(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        return MatchColumnsUsingDictionary(excelData, tableMetadata, MonthlyProgressReport_DBOSummarizeSheetLIMapping);
    }

    private List<ColumnMapping> MatchColumnsForCustomerMaster(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in CustomerMasterMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Handle composite address field (contains "+")
            if (excelColumnName.Contains("+"))
            {
                // Split by "+" to get individual Excel column names
                var addressColumns = excelColumnName.Split('+', StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if all address columns exist in Excel
                var allAddressColumnsExist = addressColumns.All(col =>
                    excelColumns.Any(ec => ec.ColumnName.Equals(col, StringComparison.OrdinalIgnoreCase)));

                // Check if SQL table has the Address column
                if (allAddressColumnsExist && sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                {
                    // Add a special mapping for composite address
                    // We'll use a special marker in ExcelColumnName to identify composite fields
                    // and handle concatenation in PrepareMappedDataTableAsync
                    mappings.Add(new ColumnMapping
                    {
                        ExcelColumnName = excelColumnName, // Store the full composite key for identification
                        SqlColumnName = sqlColumn.ColumnName,
                        SqlDataType = sqlColumn.DataType,
                        IsIdentity = sqlColumn.IsIdentity,
                        IsNullable = sqlColumn.IsNullable
                    });
                }
                continue;
            }

            // Regular single column mapping
            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumnRegular))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumnRegular.ColumnName,
                SqlDataType = sqlColumnRegular.DataType,
                IsIdentity = sqlColumnRegular.IsIdentity,
                IsNullable = sqlColumnRegular.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForCustomerContacts(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in CustomerContactMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Handle composite address field (contains "+")
            if (excelColumnName.Contains("+"))
            {
                // Split by "+" to get individual Excel column names
                var addressColumns = excelColumnName.Split('+', StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if all address columns exist in Excel
                var allAddressColumnsExist = addressColumns.All(col =>
                    excelColumns.Any(ec => ec.ColumnName.Equals(col, StringComparison.OrdinalIgnoreCase)));

                // Check if SQL table has the Address column
                if (allAddressColumnsExist && sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                {
                    // Add a special mapping for composite address
                    // We'll use a special marker in ExcelColumnName to identify composite fields
                    // and handle concatenation in PrepareMappedDataTableAsync
                    mappings.Add(new ColumnMapping
                    {
                        ExcelColumnName = excelColumnName, // Store the full composite key for identification
                        SqlColumnName = sqlColumn.ColumnName,
                        SqlDataType = sqlColumn.DataType,
                        IsIdentity = sqlColumn.IsIdentity,
                        IsNullable = sqlColumn.IsNullable
                    });
                }
                continue;
            }

            // Regular single column mapping
            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumnRegular))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumnRegular.ColumnName,
                SqlDataType = sqlColumnRegular.DataType,
                IsIdentity = sqlColumnRegular.IsIdentity,
                IsNullable = sqlColumnRegular.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForVendorMaster(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in VendorMasterMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Handle composite address field (contains "+")
            if (excelColumnName.Contains("+"))
            {
                // Split by "+" to get individual Excel column names
                var addressColumns = excelColumnName.Split('+', StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .ToList();

                // Check if all address columns exist in Excel
                var allAddressColumnsExist = addressColumns.All(col =>
                    excelColumns.Any(ec => ec.ColumnName.Equals(col, StringComparison.OrdinalIgnoreCase)));

                // Check if SQL table has the Address column
                if (allAddressColumnsExist && sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                {
                    // Add a special mapping for composite address
                    // We'll use a special marker in ExcelColumnName to identify composite fields
                    // and handle concatenation in PrepareMappedDataTableAsync
                    mappings.Add(new ColumnMapping
                    {
                        ExcelColumnName = excelColumnName, // Store the full composite key for identification
                        SqlColumnName = sqlColumn.ColumnName,
                        SqlDataType = sqlColumn.DataType,
                        IsIdentity = sqlColumn.IsIdentity,
                        IsNullable = sqlColumn.IsNullable
                    });
                }
                continue;
            }

            // Regular single column mapping
            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumnRegular))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumnRegular.ColumnName,
                SqlDataType = sqlColumnRegular.DataType,
                IsIdentity = sqlColumnRegular.IsIdentity,
                IsNullable = sqlColumnRegular.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForBPComments(DataTable excelData, List<ColumnMetadata> tableMetadata, string attachmentRecordType)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary

        foreach (var mappingEntry in BPCommentsMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Find the Excel column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column doesn't exist

            // Find the SQL column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column doesn't exist

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });

            // Special handling for parent_object_id: map to other RecordIDs based on BP ID logic
            // Since we are inside the loop iterating BPCommentsMapping, and it contains "parent_object_id" -> "OrderTransmittalRecordID"
            // we check if this is that entry, then add the others.
            //if (excelColumnName.Equals("parent_object_id", StringComparison.OrdinalIgnoreCase))
            //{
            //    // List of other potential target columns
            //    var otherTargetColumns = new[] { "TurbineRecordID", "MechanicalDBORecordID", "ElectricalInstrumentationDBORecordID", "OrderTransmittalRecordID", "BankGuaranteeRecordID" };

            //    foreach (var targetCol in otherTargetColumns)
            //    {
            //        if (sqlColumnLookup.TryGetValue(targetCol, out var targetSqlColumn))
            //        {
            //            mappings.Add(new ColumnMapping
            //            {
            //                ExcelColumnName = excelColumn.ColumnName,
            //                SqlColumnName = targetSqlColumn.ColumnName,
            //                SqlDataType = targetSqlColumn.DataType,
            //                IsIdentity = targetSqlColumn.IsIdentity,
            //                IsNullable = targetSqlColumn.IsNullable
            //            });
            //        }
            //    }
            //}
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForTurbine(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in TurbineMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Find the Excel column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column doesn't exist

            // Find the SQL column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column doesn't exist

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForBPAttachments(DataTable excelData, List<ColumnMetadata> tableMetadata, string? attachmentRecordType = null)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Select the appropriate mapping dictionary based on AttachmentRecordType
        Dictionary<string, string> selectedMapping;
        if (!string.IsNullOrWhiteSpace(attachmentRecordType))
        {
            if (string.Equals(attachmentRecordType, "Comment", StringComparison.OrdinalIgnoreCase))
            {
                selectedMapping = BPAttachmentCommentMapping;
            }
            else if (string.Equals(attachmentRecordType, "OrderTransmittal", StringComparison.OrdinalIgnoreCase))
            {
                selectedMapping = BPAttachmentOTMapping;
            }
            else
            {
                // Default to BPAttachmentMapping for unknown types
                selectedMapping = BPAttachmentMapping;
            }
        }
        else
        {
            // Default to BPAttachmentMapping if no AttachmentRecordType is provided
            selectedMapping = BPAttachmentMapping;
        }

        // Check if Excel has parent_type column (needed for conditional mapping)
        var hasParentTypeColumn = excelColumns.Any(ec =>
            ec.ColumnName.Equals("parent_type", StringComparison.OrdinalIgnoreCase));
        
        // Iterate through the selected hardcoded mapping dictionary
        foreach (var mappingEntry in selectedMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Handle file_name → FileName only.
            // FilePath is exclusively populated by node_path in the regular mapping path below.
            // Rows where node_path is NULL will insert with FilePath = NULL (no rows are skipped).
            if (excelColumnName.Equals("file_name", StringComparison.OrdinalIgnoreCase))
            {
                if (sqlColumnLookup.TryGetValue("FileName", out var fileNameColumn))
                {
                    var excelColumn = excelColumns.FirstOrDefault(
                        ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));
                    if (excelColumn != null)
                    {
                        mappings.Add(new ColumnMapping
                        {
                            ExcelColumnName = excelColumn.ColumnName,
                            SqlColumnName = fileNameColumn.ColumnName,
                            SqlDataType = fileNameColumn.DataType,
                            IsIdentity = fileNameColumn.IsIdentity,
                            IsNullable = fileNameColumn.IsNullable
                        });
                    }
                }
                continue;
            }

            // Handle parent_id - conditionally map to OrderTransmittalRecordID and other RecordIDs
            // We'll add the mapping but handle the conditional logic in PrepareMappedDataTableAsync


            // Regular single column mapping
            var excelColumnRegular = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumnRegular == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumnRegular))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumnRegular.ColumnName,
                SqlColumnName = sqlColumnRegular.ColumnName,
                SqlDataType = sqlColumnRegular.DataType,
                IsIdentity = sqlColumnRegular.IsIdentity,
                // Force FilePath to be nullable in BPAttachments: rows where node_path is NULL
                // must still be inserted (with FilePath = NULL), even if the SQL schema says NOT NULL.
                IsNullable = sqlColumnName.Equals("FilePath", StringComparison.OrdinalIgnoreCase)
                             ? true
                             : sqlColumnRegular.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForProject(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in ProjectMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForMechanicalDBO(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in MechanicalDBOMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForOrderTransmittal(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in OrderTransmittalMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForElectricalInstrumentationDBO(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in ElectricalInstrumentationDBOMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable,
                ForeignKeyTableSchema = sqlColumn.ForeignKeyTableSchema,
                ForeignKeyTableName = sqlColumn.ForeignKeyTableName,
                ForeignKeyColumnName = sqlColumn.ForeignKeyColumnName,
                ForeignKeyLookupColumnName = sqlColumn.ForeignKeyLookupColumnName
            });
        }

        return mappings;
    }

    private List<ColumnMapping> MatchColumnsForApprovalLog(DataTable excelData, List<ColumnMetadata> tableMetadata)
    {
        var mappings = new List<ColumnMapping>();
        var excelColumns = excelData.Columns.Cast<DataColumn>().ToList();

        // Create a lookup for SQL column metadata by column name (case-insensitive)
        var sqlColumnLookup = tableMetadata.ToDictionary(
            m => m.ColumnName,
            m => m,
            StringComparer.OrdinalIgnoreCase);

        // Iterate through the hardcoded mapping dictionary
        foreach (var mappingEntry in ApprovalLogColumnMapping)
        {
            var excelColumnName = mappingEntry.Key;
            var sqlColumnName = mappingEntry.Value;

            // Check if Excel has this column
            var excelColumn = excelColumns.FirstOrDefault(
                ec => ec.ColumnName.Equals(excelColumnName, StringComparison.OrdinalIgnoreCase));

            if (excelColumn == null)
                continue; // Skip if Excel column not found

            // Check if SQL table has the mapped column
            if (!sqlColumnLookup.TryGetValue(sqlColumnName, out var sqlColumn))
                continue; // Skip if SQL column not found in metadata

            // Add the mapping
            mappings.Add(new ColumnMapping
            {
                ExcelColumnName = excelColumn.ColumnName,
                SqlColumnName = sqlColumn.ColumnName,
                SqlDataType = sqlColumn.DataType,
                IsIdentity = sqlColumn.IsIdentity,
                IsNullable = sqlColumn.IsNullable
            });
        }

        return mappings;
    }
    private async Task CreateTempTableAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string tempTableName,
        List<ColumnMetadata> metadata,
        CancellationToken cancellationToken)
    {
        var sb = new StringBuilder();
        sb.Append($"CREATE TABLE {tempTableName} (");

        var columns = new List<string>();
        foreach (var col in metadata)
        {
            var columnDef = $"[{col.ColumnName}] {GetSqlTypeDefinition(col)}";
            columns.Add(columnDef);
        }

        sb.Append(string.Join(", ", columns));
        sb.Append(")");

        await using var command = new SqlCommand(sb.ToString(), connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private string GetSqlTypeDefinition(ColumnMetadata metadata)
    {
        var type = metadata.DataType.ToUpper();

        switch (type)
        {
            case "VARCHAR":
            case "NVARCHAR":
            case "CHAR":
            case "NCHAR":
                var length = metadata.MaxLength ?? 255;
                if (length == -1) length = 4000; // MAX
                return $"{type}({length})";

            case "DECIMAL":
            case "NUMERIC":
                var precision = metadata.NumericPrecision ?? 18;
                var scale = metadata.NumericScale ?? 0;
                return $"{type}({precision},{scale})";

            case "FLOAT":
                return metadata.NumericPrecision.HasValue
                    ? $"{type}({metadata.NumericPrecision.Value})"
                    : "FLOAT";

            default:
                return type;
        }
    }

    private async Task<(DataTable mappedTable, List<Models.RowErrorDetail> rowErrors)> PrepareMappedDataTableAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        DataTable excelData,
        List<ColumnMapping> mappings,
        string tableName,
        string schemaName,
        List<ColumnMetadata> tableMetadata,
        CancellationToken cancellationToken)
    {
        var mappedTable = new DataTable();
        var rowErrors = new List<Models.RowErrorDetail>();
        var isCommunicationProtocol = string.Equals(tableName, "CommunicationProtocol", StringComparison.OrdinalIgnoreCase);
        var isBankGuarantee = string.Equals(tableName, "BankGuarantee", StringComparison.OrdinalIgnoreCase);
        var isCustomerMaster = string.Equals(tableName, "CustomerMaster", StringComparison.OrdinalIgnoreCase);
        var isCustomerContacts = string.Equals(tableName, "CustomerContacts", StringComparison.OrdinalIgnoreCase);
        var isVendorMaster = string.Equals(tableName, "VendorMaster", StringComparison.OrdinalIgnoreCase);
        var isBPAttachments = string.Equals(tableName, "BPAttachments", StringComparison.OrdinalIgnoreCase);
        var isBPComments = string.Equals(tableName, "BPComments", StringComparison.OrdinalIgnoreCase);
        var isTurbine = tableName.StartsWith("Turbine", StringComparison.OrdinalIgnoreCase);
        var isProject = string.Equals(tableName, "Project", StringComparison.OrdinalIgnoreCase);
        var isMechanicalDBO = tableName.StartsWith("MechanicalDBO", StringComparison.OrdinalIgnoreCase);
        var isOrderTransmittal = tableName.StartsWith("OrderTransmittal", StringComparison.OrdinalIgnoreCase);
        var isElectricalInstrumentationDBO = tableName.StartsWith("ElectricalInstrumentationDBO", StringComparison.OrdinalIgnoreCase);
        var isLetterOfCorrespondence = string.Equals(tableName, "LetterOfCorrespondence", StringComparison.OrdinalIgnoreCase);
        var isContractClearance = string.Equals(tableName, "ContractClearance", StringComparison.OrdinalIgnoreCase);
        var isAdditionalOrderBooking = string.Equals(tableName, "AdditionalOrderBooking", StringComparison.OrdinalIgnoreCase);
        var isContractOnHold = tableName.StartsWith("ContractOnHold", StringComparison.OrdinalIgnoreCase);
        var isMinutesOfMeeting = tableName.StartsWith("MinutesOfMeeting", StringComparison.OrdinalIgnoreCase) || 
                               tableName.StartsWith("MOM", StringComparison.OrdinalIgnoreCase);
        var isLCReview = tableName.StartsWith("LCReview", StringComparison.OrdinalIgnoreCase);
        var isLCReviewNotesObservation = string.Equals(tableName, "LCReview_NotesObservation", StringComparison.OrdinalIgnoreCase);
        var isInitialCashPlan = string.Equals(tableName, "InitialCashPlan", StringComparison.OrdinalIgnoreCase);
        var isPaymentSupply = string.Equals(tableName, "payment_supply", StringComparison.OrdinalIgnoreCase);
        var isLiquidatedDamage = string.Equals(tableName, "LiquidatedDamage", StringComparison.OrdinalIgnoreCase) || string.Equals(tableName, "liquidated_damage", StringComparison.OrdinalIgnoreCase);
        var isPaymentENC = string.Equals(tableName, "payment_enc", StringComparison.OrdinalIgnoreCase);
        var isMonthlyActualCollection = string.Equals(tableName, "MonthlyActualCollection", StringComparison.OrdinalIgnoreCase);
        var isInitialCashFlowPlan = string.Equals(tableName, "InitialCashFlowPlan", StringComparison.OrdinalIgnoreCase);
        var isMonthlyPlanning = string.Equals(tableName, "MonthlyPlanning", StringComparison.OrdinalIgnoreCase);
        var isMonthlyPlanningLineItem = string.Equals(tableName, "MonthlyPlanningLineItem", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(tableName, "MonthlyPlanning_LineItem", StringComparison.OrdinalIgnoreCase);
        var isMonthlyActualCollectionPlanned = string.Equals(tableName, "MonthlyActualCollectionPlanned", StringComparison.OrdinalIgnoreCase);
        var isMonthlyActualUnplannedCollection = string.Equals(tableName, "MonthlyActualUnplannedCollection", StringComparison.OrdinalIgnoreCase);
        var isSpecificationRelease = string.Equals(tableName, "SpecificationRelease", StringComparison.OrdinalIgnoreCase);
        var isSparesOrderTransmittal = string.Equals(tableName, "SparesOrderTransmittal", StringComparison.OrdinalIgnoreCase);
        var isSparesOrderTransmittalLineItem = string.Equals(tableName, "SparesOrderTransmittalLineItem", StringComparison.OrdinalIgnoreCase);
        var isOrderReceiptAcknowledgement = string.Equals(tableName, "OrderReceiptAcknowledgement", StringComparison.OrdinalIgnoreCase);
        var isAuditAction = string.Equals(tableName, "AuditAction", StringComparison.OrdinalIgnoreCase);
        var isRCCA = string.Equals(tableName, "RCCA", StringComparison.OrdinalIgnoreCase);
        var isOTRisk = string.Equals(tableName, "OTRisk", StringComparison.OrdinalIgnoreCase);
        var isMonthlyProgressReport = tableName.StartsWith("MonthlyProgressReport", StringComparison.OrdinalIgnoreCase);
        var isOrderTransmittalNotes = string.Equals(tableName, "OrderTransmittal_Notes", StringComparison.OrdinalIgnoreCase);

        // Performance optimization: Cache FK lookups to avoid repeated database queries
        var projectIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var unitIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var customerIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var orderTransmittalIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var projectTypeMasterIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var momIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var lcReviewIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var initialCashPlanIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var monthlyPlanningIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var initialCashFlowPlanIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var monthlyPlanningLineItemIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var monthlyActualCollectionIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var monthlyProgressReportIdCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);


        // Add columns in the order of mappings
        foreach (var mapping in mappings)
        {
            // Handle composite address field (contains "+")
            if (mapping.ExcelColumnName.Contains("+"))
            {
                // For composite fields, determine type from SQL metadata
                var targetNetType = GetNetTypeFromSqlType(mapping.SqlDataType);
                if (targetNetType == null)
                {
                    targetNetType = typeof(string); // Default to string for composite address
                }

                var newColumn = new DataColumn(mapping.SqlColumnName, targetNetType)
                {
                    AllowDBNull = mapping.IsNullable
                };
                mappedTable.Columns.Add(newColumn);
                continue;
            }

            var excelColumn = excelData.Columns[mapping.ExcelColumnName];
            if (excelColumn == null)
                continue; // Skip if column not found (should not happen due to matching)

            // Determine the target .NET type based on SQL Server data type
            var targetNetTypeRegular = GetNetTypeFromSqlType(mapping.SqlDataType);
            if (targetNetTypeRegular == null)
            {
                // Fallback to Excel column type if SQL type mapping fails
                targetNetTypeRegular = excelColumn.DataType;
            }

            var newColumnRegular = new DataColumn(mapping.SqlColumnName, targetNetTypeRegular)
            {
                AllowDBNull = mapping.IsNullable
            };
            mappedTable.Columns.Add(newColumnRegular);
        }

        // Check if IsDeleted column exists in SQL table but not in mappings - add it if needed
        var isDeletedColumn = tableMetadata?.FirstOrDefault(m =>
            string.Equals(m.ColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase));
        if (isDeletedColumn != null && !mappings.Any(m =>
            string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase)))
        {
            var isDeletedNetType = GetNetTypeFromSqlType(isDeletedColumn.DataType);
            if (isDeletedNetType == null)
            {
                isDeletedNetType = typeof(bool); // Default to bool for BIT type
            }
            var isDeletedDataColumn = new DataColumn("IsDeleted", isDeletedNetType)
            {
                AllowDBNull = isDeletedColumn.IsNullable
            };
            mappedTable.Columns.Add(isDeletedDataColumn);
        }

        // Copy data
        int rowNumber = 1; // Excel row number (1-based, including header)
        foreach (DataRow excelRow in excelData.Rows)
        {
            rowNumber++; // Increment for data rows (header is row 1)
            var newRow = mappedTable.NewRow();

            // Set IsDeleted to false immediately if the column exists (before processing other columns)
            if (mappedTable.Columns.Contains("IsDeleted"))
            {
                var isDeletedDataColumn = mappedTable.Columns["IsDeleted"];
                if (isDeletedDataColumn != null && isDeletedDataColumn.DataType == typeof(bool))
                {
                    newRow["IsDeleted"] = false;
                }
                else if (isDeletedDataColumn != null)
                {
                    // For other types (int, bit as int, etc.), set to 0
                    newRow["IsDeleted"] = Convert.ChangeType(0, isDeletedDataColumn.DataType);
                }
            }

            bool skipRow = false;
            string? errorColumn = null;
            object? errorValue = null;
            string? errorMessage = null;
            var rowData = new Dictionary<string, object?>();

            // Collect all row data for error reporting
            foreach (var mapping in mappings)
            {
                try
                {
                    // Handle composite address field (contains "+")
                    if (mapping.ExcelColumnName.Contains("+"))
                    {
                        // For composite fields, collect all component columns
                        var addressColumns = mapping.ExcelColumnName.Split('+', StringSplitOptions.RemoveEmptyEntries)
                            .Select(c => c.Trim())
                            .ToList();
                        foreach (var addrCol in addressColumns)
                        {
                            if (excelData.Columns.Contains(addrCol))
                            {
                                rowData[addrCol] = excelRow[addrCol];
                            }
                        }
                    }
                    else
                    {
                        var value = excelRow[mapping.ExcelColumnName];
                        rowData[mapping.ExcelColumnName] = value;
                    }
                }
                catch
                {
                    // Ignore errors when collecting row data
                }
            }

            foreach (var mapping in mappings)
            {
                try
                {
                    object? value;

                    // Handle composite address field (contains "+")
                    if (mapping.ExcelColumnName.Contains("+"))
                    {
                        // Split by "+" to get individual Excel column names
                        var addressColumns = mapping.ExcelColumnName.Split('+', StringSplitOptions.RemoveEmptyEntries)
                            .Select(c => c.Trim())
                            .ToList();

                        // Concatenate address columns with space separator
                        var addressParts = new List<string>();
                        foreach (var addrCol in addressColumns)
                        {
                            if (excelData.Columns.Contains(addrCol))
                            {
                                var addrValue = excelRow[addrCol];
                                if (addrValue != null && addrValue != DBNull.Value)
                                {
                                    var addrStr = addrValue.ToString()?.Trim();
                                    if (!string.IsNullOrWhiteSpace(addrStr))
                                    {
                                        addressParts.Add(addrStr);
                                    }
                                }
                            }
                        }
                        value = string.Join(" ", addressParts);
                        if (string.IsNullOrWhiteSpace(value?.ToString()))
                        {
                            value = DBNull.Value;
                        }
                    }
                    else
                    {
                        value = excelRow[mapping.ExcelColumnName];
                    }

                    // Remove double quotes if present at the beginning and end, and leading single quote
                    if (value is string fieldValueAsStr && !string.IsNullOrEmpty(fieldValueAsStr))
                    {
                        if (fieldValueAsStr.StartsWith("\"") && fieldValueAsStr.EndsWith("\"") && fieldValueAsStr.Length >= 2)
                        {
                            fieldValueAsStr = fieldValueAsStr.Trim('\"');
                        }

                        if (fieldValueAsStr.StartsWith("'"))
                        {
                            fieldValueAsStr = fieldValueAsStr.TrimStart('\'');
                        }

                        value = fieldValueAsStr;
                    }



                    // Special handling for ProjectID column (FK to master.Project)
                    // Validate ProjectID exists in master.Project table
                    // Skip validation if table is "Project" itself (ProjectID is primary key, not FK)
                    if (!isProject &&
                        string.Equals(mapping.SqlColumnName, "ProjectID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        // If value is numeric 0, convert to NULL
                        if ((value is int intVal && intVal == 0) ||
                            (value is long longVal && longVal == 0) ||
                            (value is short shortVal && shortVal == 0) ||
                            (long.TryParse(value.ToString()?.Trim(), out var projectNumeric) && projectNumeric == 0))
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            // ElectricalInstrumentationDBO fast path: numeric IDs are already target keys.
                            // Avoid per-row lookup query in large uploads.
                            if (isElectricalInstrumentationDBO &&
                                long.TryParse(value.ToString()?.Trim(), out var directProjectId) &&
                                directProjectId > 0)
                            {
                                value = directProjectId;
                            }
                            else
                            {
                                var valueKey = value.ToString()?.Trim() ?? string.Empty;

                                // Check cache first
                                if (!projectIdCache.TryGetValue(valueKey, out var resolvedProjectId))
                                {
                                    // Not in cache, resolve from database
                                    resolvedProjectId = await ResolveProjectIdAsync(
                                        connection,
                                        transaction,
                                        value,
                                        cancellationToken);

                                    // Cache the result (even if null)
                                    projectIdCache[valueKey] = resolvedProjectId;
                                }

                                if (resolvedProjectId == null)
                                {
                                    // For CommunicationProtocol, OrderTransmittal, BankGuarantee, Turbine, and ElectricalInstrumentationDBO, if ProjectID doesn't exist, set to NULL instead of skipping
                                    if (isCommunicationProtocol || isOrderTransmittal || isBankGuarantee || isTurbine || isElectricalInstrumentationDBO || isBPAttachments || isMechanicalDBO || isContractClearance || isAdditionalOrderBooking || isMinutesOfMeeting || isContractOnHold || isLCReview || isInitialCashPlan || isPaymentSupply || isLiquidatedDamage || isPaymentENC || isInitialCashFlowPlan || isMonthlyPlanning || isMonthlyPlanningLineItem || isMonthlyActualCollectionPlanned || isMonthlyActualUnplannedCollection || isSpecificationRelease || isOrderReceiptAcknowledgement || isSparesOrderTransmittal || isAuditAction || isRCCA || isMonthlyProgressReport || isLetterOfCorrespondence)
                                    {
                                        value = DBNull.Value;
                                    }
                                    else
                                    {
                                        // For other tables, skip row if ProjectID doesn't exist
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: ProjectID '{value}' does not exist in table 'master.Project'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    value = resolvedProjectId;
                                }
                            }
                        }
                    }

                    // Special handling for CloneProjectId in MechanicalDBO (FK to master.Project)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CloneProjectId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                         // Reuse the ProjectID resolution logic
                         var valueKey = value.ToString()?.Trim() ?? string.Empty;
                         // Check cache first (using projectIdCache since it's the same target table)
                         if (!projectIdCache.TryGetValue(valueKey, out var resolvedProjectId))
                         {
                             resolvedProjectId = await ResolveProjectIdAsync(connection, transaction, value, cancellationToken);
                             projectIdCache[valueKey] = resolvedProjectId;
                         }

                         if (resolvedProjectId == null)
                         {
                             // If referenced project is missing, set to NULL (optional behavior, preventing failure)
                             value = DBNull.Value;
                         }
                         else
                         {
                             value = resolvedProjectId;
                         }
                    }

                    // Special handling for MOMID column in MOM child tables (FK to bp.MOM)
                    if (isMinutesOfMeeting &&
                        string.Equals(mapping.SqlColumnName, "MOMID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null &&
                        (string.Equals(tableName, "MOM_Minutes", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(tableName, "MOM_Attendees", StringComparison.OrdinalIgnoreCase)))
                    {
                        var momValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!momIdCache.TryGetValue(momValueKey, out var momIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "MOM", "MOMID", momValueKey, cancellationToken);
                            momIdExists = exists;
                            momIdCache[momValueKey] = momIdExists;
                        }

                        if (!(bool)momIdExists!)
                        {
                            // If MOM record doesn't exist, set to NULL as requested
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for LCReviewId column in LCReview child tables (FK to bp.LCReview)
                    if (isLCReviewNotesObservation &&
                        string.Equals(mapping.SqlColumnName, "LCReviewId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var lcReviewValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!lcReviewIdCache.TryGetValue(lcReviewValueKey, out var lcReviewIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "LCReview", "LCReviewId", lcReviewValueKey, cancellationToken);
                            lcReviewIdExists = exists;
                            lcReviewIdCache[lcReviewValueKey] = lcReviewIdExists;
                        }

                        if (!(bool)lcReviewIdExists!)
                        {
                            // If LCReview record doesn't exist, set to NULL (preventing failure)
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for OrderTransmittalId column in InitialCashFlowPlan (FK to bp.OrderTransmittal)
                    if ((isInitialCashFlowPlan || isRCCA || isMonthlyProgressReport) &&
                        string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var otValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!orderTransmittalIdCache.TryGetValue(otValueKey, out var orderTransmittalIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "OrderTransmittal", "OrderTransmittalID", otValueKey, cancellationToken);
                            orderTransmittalIdExists = exists;
                            orderTransmittalIdCache[otValueKey] = orderTransmittalIdExists;
                        }

                        if (!(bool)orderTransmittalIdExists!)
                        {
                            // If OrderTransmittal record doesn't exist, set to NULL (preventing failure)
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for InitialCashPlanId column in InitialCashFlowPlan (FK to bp.InitialCashPlan)
                    if ((isInitialCashFlowPlan || isMonthlyProgressReport) &&
                        string.Equals(mapping.SqlColumnName, "InitialCashPlanId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var icpValueKey = value.ToString()?.Trim() ?? string.Empty;
                        
                        if (!initialCashPlanIdCache.TryGetValue(icpValueKey, out var initialCashPlanIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "InitialCashPlan", "InitialCashPlanId", icpValueKey, cancellationToken);
                            initialCashPlanIdExists = exists;
                            initialCashPlanIdCache[icpValueKey] = initialCashPlanIdExists;
                        }

                        if (!(bool)initialCashPlanIdExists!)
                        {
                            // If InitialCashPlan record doesn't exist, set to NULL (preventing failure)
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for MonthlyPlanningId column in MonthlyPlanningLineItem (FK to bp.MonthlyPlanning)
                    if (isMonthlyPlanningLineItem &&
                        string.Equals(mapping.SqlColumnName, "MonthlyPlanningId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var mpValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!monthlyPlanningIdCache.TryGetValue(mpValueKey, out var monthlyPlanningIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "MonthlyPlanning", "MonthlyPlanningId", mpValueKey, cancellationToken);
                            monthlyPlanningIdExists = exists;
                            monthlyPlanningIdCache[mpValueKey] = monthlyPlanningIdExists;
                        }

                        if (!(bool)monthlyPlanningIdExists!)
                        {
                            // If MonthlyPlanning record doesn't exist, set to NULL
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for InitialCashPlanId column in MonthlyPlanningLineItem (FK to bp.InitialCashPlan)
                    if (isMonthlyPlanningLineItem &&
                        string.Equals(mapping.SqlColumnName, "InitialCashPlanId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var icpValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!initialCashPlanIdCache.TryGetValue(icpValueKey, out var initialCashPlanIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "InitialCashPlan", "InitialCashPlanId", icpValueKey, cancellationToken);
                            initialCashPlanIdExists = exists;
                            initialCashPlanIdCache[icpValueKey] = initialCashPlanIdExists;
                        }

                        if (!(bool)initialCashPlanIdExists!)
                        {
                            // If InitialCashPlan record doesn't exist, set to NULL
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for InitialCashFlowPlanId column in MonthlyPlanningLineItem (FK to bp.InitialCashFlowPlan)
                    if ((isMonthlyPlanningLineItem || isMonthlyProgressReport) &&
                        string.Equals(mapping.SqlColumnName, "InitialCashFlowPlanId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var icfpValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!initialCashFlowPlanIdCache.TryGetValue(icfpValueKey, out var initialCashFlowPlanIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "InitialCashFlowPlan", "InitialCashFlowPlanId", icfpValueKey, cancellationToken);
                            initialCashFlowPlanIdExists = exists;
                            initialCashFlowPlanIdCache[icfpValueKey] = initialCashFlowPlanIdExists;
                        }

                        if (!(bool)initialCashFlowPlanIdExists!)
                        {
                            // If InitialCashFlowPlan record doesn't exist, set to NULL
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for OrderTransmittalId column in MonthlyPlanningLineItem (FK to bp.OrderTransmittal)
                    if (isMonthlyPlanningLineItem &&
                        string.Equals(mapping.SqlColumnName, "OrderTransmittalId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var otValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!orderTransmittalIdCache.TryGetValue(otValueKey, out var orderTransmittalIdExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "OrderTransmittal", "OrderTransmittalID", otValueKey, cancellationToken);
                            orderTransmittalIdExists = exists;
                            orderTransmittalIdCache[otValueKey] = orderTransmittalIdExists;
                        }

                        if (!(bool)orderTransmittalIdExists!)
                        {
                            // If OrderTransmittal record doesn't exist, set to NULL
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for MonthlyPlanningId column in MonthlyActualCollectionPlanned (FK to bp.MonthlyPlanning)
                    if (isMonthlyActualCollectionPlanned &&
                        string.Equals(mapping.SqlColumnName, "MonthlyPlanningId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var mpValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(mpValueKey, out var mpD)) { mpValueKey = ((long)mpD).ToString(); value = (long)mpD; }

                        if (!string.IsNullOrWhiteSpace(mpValueKey))
                        {
                            if (!monthlyPlanningIdCache.TryGetValue(mpValueKey, out var mpExists))
                            {
                                mpExists = await RecordExistsAsync(connection, transaction, schemaName, "MonthlyPlanning", "MonthlyPlanningId", mpValueKey, cancellationToken);
                                monthlyPlanningIdCache[mpValueKey] = mpExists;
                            }
                            if (!(bool)mpExists!) value = DBNull.Value;
                        }
                    }

                    // Special handling for MonthlyPlanningLineItemId column in MonthlyActualCollectionPlanned (FK to bp.MonthlyPlanningLineItem)
                    if (isMonthlyActualCollectionPlanned &&
                        string.Equals(mapping.SqlColumnName, "MonthlyPlanningLineItemId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var mpliValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(mpliValueKey, out var mpliD)) { mpliValueKey = ((long)mpliD).ToString(); value = (long)mpliD; }

                        if (!string.IsNullOrWhiteSpace(mpliValueKey))
                        {
                            if (!monthlyPlanningLineItemIdCache.TryGetValue(mpliValueKey, out var mpliExists))
                            {
                                mpliExists = await RecordExistsAsync(connection, transaction, schemaName, "MonthlyPlanning_LineItem", "MonthlyPlanningLineItemId", mpliValueKey, cancellationToken);
                                monthlyPlanningLineItemIdCache[mpliValueKey] = mpliExists;
                            }
                            if (!(bool)mpliExists!) value = DBNull.Value;
                        }
                    }

                    // Special handling for OrderTransmittalID column in MonthlyActualCollectionPlanned (FK to bp.OrderTransmittal)
                    if (isMonthlyActualCollectionPlanned &&
                        string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var otValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(otValueKey, out var otD)) { otValueKey = ((long)otD).ToString(); value = (long)otD; }

                        if (!string.IsNullOrWhiteSpace(otValueKey))
                        {
                            if (!orderTransmittalIdCache.TryGetValue(otValueKey, out var otExists))
                            {
                                otExists = await RecordExistsAsync(connection, transaction, schemaName, "OrderTransmittal", "OrderTransmittalID", otValueKey, cancellationToken);
                                orderTransmittalIdCache[otValueKey] = otExists;
                            }
                            if (!(bool)otExists!) value = DBNull.Value;
                        }
                    }

                    // Special handling for MonthlyProgressReportID column in MonthlyProgressReport line items (FK to bp.MonthlyProgressReport)
                    if (isMonthlyProgressReport &&
                        string.Equals(mapping.SqlColumnName, "MonthlyProgressReportID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var mprValueKey = value.ToString()?.Trim() ?? string.Empty;

                        if (!monthlyProgressReportIdCache.TryGetValue(mprValueKey, out var mprExists))
                        {
                            var exists = await RecordExistsAsync(connection, transaction, "bp", "MonthlyProgressReport", "MonthlyProgressReportID", mprValueKey, cancellationToken);
                            mprExists = exists;
                            monthlyProgressReportIdCache[mprValueKey] = mprExists;
                        }

                        if (!(bool)mprExists!)
                        {
                            // If parent record doesn't exist, set to NULL
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for MonthlyActualCollectionId column in MonthlyActualCollectionPlanned (FK to bp.MonthlyActualCollection)
                    if (isMonthlyActualCollectionPlanned &&
                        string.Equals(mapping.SqlColumnName, "MonthlyActualCollectionId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var macValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(macValueKey, out var macD)) { macValueKey = ((long)macD).ToString(); value = (long)macD; }

                        if (!string.IsNullOrWhiteSpace(macValueKey))
                        {
                            if (!monthlyActualCollectionIdCache.TryGetValue(macValueKey, out var macExists))
                            {
                                macExists = await RecordExistsAsync(connection, transaction, schemaName, "MonthlyActualCollection", "MonthlyActualCollectionId", macValueKey, cancellationToken);
                                monthlyActualCollectionIdCache[macValueKey] = macExists;
                            }
                            if (!(bool)macExists!) value = DBNull.Value;
                        }
                    }

                    if (isSparesOrderTransmittalLineItem &&
                        string.Equals(mapping.SqlColumnName, "SparesOrderTransmittalID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        // If value is numeric 0, convert to NULL to avoid FK constraint error
                        if ((value is int intVal && intVal == 0) ||
                            (value is long longVal && longVal == 0) ||
                            (value is short shortVal && shortVal == 0) ||
                            (long.TryParse(value.ToString()?.Trim(), out var numericVal) && numericVal == 0))
                        {
                            value = DBNull.Value;
                        }
                    }


                    // Special handling for InitialCashPlanId column in MonthlyActualUnplannedCollection (FK to bp.InitialCashPlan)
                    if (isMonthlyActualUnplannedCollection &&
                        string.Equals(mapping.SqlColumnName, "InitialCashPlanId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var icpValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(icpValueKey, out var icpD)) { icpValueKey = ((long)icpD).ToString(); value = (long)icpD; }

                        if (!string.IsNullOrWhiteSpace(icpValueKey))
                        {
                            if (!initialCashPlanIdCache.TryGetValue(icpValueKey, out var icpExists))
                            {
                                icpExists = await RecordExistsAsync(connection, transaction, schemaName, "InitialCashPlan", "InitialCashPlanId", icpValueKey, cancellationToken);
                                initialCashPlanIdCache[icpValueKey] = icpExists;
                            }
                            if (!(bool)icpExists!) value = DBNull.Value;
                        }
                    }

                    // Special handling for OrdertransmittalId column in MonthlyActualUnplannedCollection (FK to bp.OrderTransmittal)
                    if (isMonthlyActualUnplannedCollection &&
                        string.Equals(mapping.SqlColumnName, "OrdertransmittalId", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var otValueKey = value.ToString()?.Trim() ?? string.Empty;
                        if (double.TryParse(otValueKey, out var otD2)) { otValueKey = ((long)otD2).ToString(); value = (long)otD2; }

                        if (!string.IsNullOrWhiteSpace(otValueKey))
                        {
                            if (!orderTransmittalIdCache.TryGetValue(otValueKey, out var otExists2))
                            {
                                otExists2 = await RecordExistsAsync(connection, transaction, schemaName, "OrderTransmittal", "OrderTransmittalID", otValueKey, cancellationToken);
                                orderTransmittalIdCache[otValueKey] = otExists2;
                            }
                            if (!(bool)otExists2!) value = DBNull.Value;
                        }
                    }

                    // Special handling for ProjectTypeMasterID column (FK to master.ProjectTypeMaster)
                    // Resolve by ProjectTypeMasterID (numeric) or by ProjectTypeName (string)
                    if (string.Equals(mapping.SqlColumnName, "ProjectTypeMasterID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var valueStr = value.ToString()?.Trim() ?? string.Empty;
                        if (string.IsNullOrEmpty(valueStr) || (long.TryParse(valueStr, out var projectTypeNumeric) && projectTypeNumeric == 0))
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = valueStr;

                            // Check cache first
                            if (!projectTypeMasterIdCache.TryGetValue(valueKey, out var resolvedProjectTypeMasterId))
                            {
                                // Not in cache, resolve from database
                                resolvedProjectTypeMasterId = await ResolveProjectTypeMasterIdByNameAsync(
                                    connection,
                                    transaction,
                                    value,
                                    "master",
                                    "ProjectTypeMaster",
                                    "ProjectTypeName",
                                    cancellationToken);

                                // Cache the result (even if null)
                                projectTypeMasterIdCache[valueKey] = resolvedProjectTypeMasterId;
                            }

                            if (resolvedProjectTypeMasterId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: ProjectTypeMasterID '{value}' does not exist in table 'master.ProjectTypeMaster'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedProjectTypeMasterId;
                        }
                    }

                    // Special handling for CustomerMasterID and EndUserID in ContractClearance (FK to master.CustomerMaster) 
                    // Special handling for CustomerId in CommunicationProtocol (FK to master.CustomerMaster)
                    // If the customer doesn't exist, set to NULL instead of failing the migration
                    if ( ((isContractClearance && (string.Equals(mapping.SqlColumnName, "CustomerMasterID", StringComparison.OrdinalIgnoreCase) || string.Equals(mapping.SqlColumnName, "EndUserID", StringComparison.OrdinalIgnoreCase))) ||
                          (isCommunicationProtocol && string.Equals(mapping.SqlColumnName, "CustomerId", StringComparison.OrdinalIgnoreCase)) ||
                          (isLetterOfCorrespondence && (string.Equals(mapping.SqlColumnName, "CustomerMasterID", StringComparison.OrdinalIgnoreCase) || string.Equals(mapping.SqlColumnName, "EndUserID", StringComparison.OrdinalIgnoreCase))) ||
                          (isSparesOrderTransmittal && (string.Equals(mapping.SqlColumnName, "CustomerID", StringComparison.OrdinalIgnoreCase) || string.Equals(mapping.SqlColumnName, "EndUserID", StringComparison.OrdinalIgnoreCase)))) &&
                         value != DBNull.Value && value != null )
                    {
                        var valueKey = value.ToString()?.Trim() ?? string.Empty;

                        // Check cache first
                        if (!customerIdCache.TryGetValue(valueKey, out var resolvedCustomerId))
                        {
                            // Not in cache, resolve from database
                            resolvedCustomerId = await ResolveCustomerIdByNameAsync(
                                connection,
                                transaction,
                                value,
                                cancellationToken);

                            // Cache the result (even if null)
                            customerIdCache[valueKey] = resolvedCustomerId;
                        }

                        if (resolvedCustomerId == null)
                        {
                            // If Customer doesn't exist, set to NULL instead of skipping row/failing
                            value = DBNull.Value;
                        }
                        else
                        {
                            value = resolvedCustomerId;
                        }
                    }

                    // Special handling for CCRecordSelectionId in ContractClearance (from k__cc_sel_rec_dpk)
                    // If the value is 0, set to NULL
                    if (isContractClearance &&
                        string.Equals(mapping.SqlColumnName, "CCRecordSelectionId", StringComparison.OrdinalIgnoreCase) &&
                        IsZeroValue(value))
                    {
                        value = DBNull.Value;
                    }

                    // Special handling for OrderTransmittalId in ContractClearance or AdditionalOrderBooking (FK to bp.OrderTransmittal)
                    // If the OrderTransmittal doesn't exist, set to NULL instead of failing the migration
                    if ((isContractClearance || isAdditionalOrderBooking || isContractOnHold || isLCReview || isInitialCashPlan || isPaymentSupply || isLiquidatedDamage || isPaymentENC || isSpecificationRelease || isOrderReceiptAcknowledgement || isLetterOfCorrespondence) &&
                        string.Equals(mapping.SqlColumnName, "OrderTransmittalId", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsValueZero(value))
                        {
                            value = DBNull.Value;
                        }
                        else if (value != DBNull.Value && value != null)
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!orderTransmittalIdCache.TryGetValue(valueKey, out var resolvedOrderTransmittalId))
                            {
                                // Not in cache, resolve from database
                                resolvedOrderTransmittalId = await ResolveOrderTransmittalIdAsync(
                                    connection,
                                    transaction,
                                    value,
                                    cancellationToken);

                                // Cache the result (even if null)
                                orderTransmittalIdCache[valueKey] = resolvedOrderTransmittalId;
                            }

                            if (resolvedOrderTransmittalId == null)
                            {
                                // If OrderTransmittal doesn't exist, set to NULL instead of failing
                                value = DBNull.Value;
                            }
                            else
                            {
                                value = resolvedOrderTransmittalId;
                            }
                        }
                    }

                    var targetColumn = mappedTable.Columns[mapping.SqlColumnName];

                    // Special handling for status column in CommunicationProtocol
                    if (isCommunicationProtocol &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    if (isMinutesOfMeeting &&
                        string.Equals(mapping.SqlColumnName, "status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for IsPresent column in MOM Attendees
                    if (isMinutesOfMeeting &&
                        string.Equals(mapping.SqlColumnName, "IsPresent", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMOMIsPresentValue(value, mapping.IsNullable);
                    }

                    if (isMinutesOfMeeting &&
                        string.Equals(mapping.SqlColumnName, "MeetingType", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMOMMeetingTypeValue(value, mapping.IsNullable);
                    }




                    if (isPaymentSupply &&
                        string.Equals(mapping.SqlColumnName, "TypeOfPayment", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var strValue = value.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(strValue, "Advance", StringComparison.OrdinalIgnoreCase))
                            value = 2;
                        else if (string.Equals(strValue, "Final", StringComparison.OrdinalIgnoreCase))
                            value = 1;
                        else if (long.TryParse(strValue, out var longVal))
                            value = longVal;
                    }

                    if (isPaymentENC &&
                        string.Equals(mapping.SqlColumnName, "TypeOfPayment", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var strValue = value.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(strValue, "Milestone Advance", StringComparison.OrdinalIgnoreCase))
                            value = 1;
                        else if (string.Equals(strValue, "Token", StringComparison.OrdinalIgnoreCase))
                            value = 2;
                        else if (string.Equals(strValue, "Final", StringComparison.OrdinalIgnoreCase))
                            value = 3;
                        else if (long.TryParse(strValue, out var longVal))
                            value = longVal;
                    }

                    if (isLiquidatedDamage &&
                        string.Equals(mapping.SqlColumnName, "LiquidatedDamageType", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        var strValue = value.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(strValue, "Late Delivery", StringComparison.OrdinalIgnoreCase)) value = 1;
                        else if (string.Equals(strValue, "Reduced Power", StringComparison.OrdinalIgnoreCase)) value = 2;
                        else if (string.Equals(strValue, "Increase in Auxiliary Power", StringComparison.OrdinalIgnoreCase)) value = 3;
                        else if (string.Equals(strValue, "Increase in Steam Consumption", StringComparison.OrdinalIgnoreCase)) value = 4;
                        else if (string.Equals(strValue, "Heat Rate", StringComparison.OrdinalIgnoreCase)) value = 5;
                        else if (string.Equals(strValue, "Documentation", StringComparison.OrdinalIgnoreCase)) value = 6;
                        else if (string.Equals(strValue, "Minimum Availability", StringComparison.OrdinalIgnoreCase)) value = 7;
                        else if (string.Equals(strValue, "Cumulative Liquidated Damages", StringComparison.OrdinalIgnoreCase)) value = 8;
                        else if (string.Equals(strValue, "Others", StringComparison.OrdinalIgnoreCase)) value = 9;
                        else if (long.TryParse(strValue, out var longVal)) value = longVal;
                    }

                    if (isLiquidatedDamage &&
                        (string.Equals(mapping.SqlColumnName, "IsAmountPercent_Maximum", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "IsAmountPercent_Minimum", StringComparison.OrdinalIgnoreCase)) &&
                        value != DBNull.Value && value != null)
                    {
                        var strValue = value.ToString()?.Trim() ?? string.Empty;
                        if (string.Equals(strValue, "Amount", StringComparison.OrdinalIgnoreCase)) value = 1;
                        else if (string.Equals(strValue, "Percentage", StringComparison.OrdinalIgnoreCase)) value = 2;
                        else if (long.TryParse(strValue, out var longVal)) value = longVal;
                    }

                    // Special handling for status column in CustomerMaster
                    if (isCustomerMaster &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for status column in CustomerContacts
                    if (isCustomerContacts &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for status column in VendorMaster
                    if (isVendorMaster &&
                        string.Equals(mapping.SqlColumnName, "StatusID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for status column in Project
                    if (isProject &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformProjectStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for ProjectTemplateID column in Project
                    if (isProject &&
                        string.Equals(mapping.SqlColumnName, "ProjectTemplateID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformProjectTemplateIdValue(value, mapping.IsNullable);
                    }

                    // Special handling for status column in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for scope columns in MechanicalDBO (TTL→0, Customer→1, Existing→2, Not Applicable→3)
                    if (isMechanicalDBO &&
                        (string.Equals(mapping.SqlColumnName, "AdditionalBOPScope", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "CondenserScope", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "GlandVentCondenserScope", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "CondensateExtractionPumpScope", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "EjectionSystemScope", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "MSParameterGlandSealingEjectionSystemScope", StringComparison.OrdinalIgnoreCase)) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOScopeValue(value, mapping.IsNullable);
                    }

                    // Special handling for Type column in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Type", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOTypeValue(value, mapping.IsNullable);
                    }

                    // Special handling for TemperatureUnitID and AmbientTemperatureUnitID in MechanicalDBO (resolve by unit name)
                    if (isMechanicalDBO &&
                        (string.Equals(mapping.SqlColumnName, "TemperatureUnitID", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "AmbientTemperatureUnitID", StringComparison.OrdinalIgnoreCase)) &&
                        value != DBNull.Value && value != null)
                    {
                        if (long.TryParse(value.ToString()?.Trim(), out var unitNumeric) && unitNumeric == 0)
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!unitIdCache.TryGetValue(valueKey, out var resolvedUnitId))
                            {
                                // Not in cache, resolve from database
                                resolvedUnitId = await ResolveUnitIdByNameAsync(
                                    connection,
                                    transaction,
                                    value,
                                    "master",
                                    "UnitMaster",
                                    "UnitName",
                                    cancellationToken);

                                // Cache the result (even if null)
                                unitIdCache[valueKey] = resolvedUnitId;
                            }

                            if (resolvedUnitId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: Unit '{value}' does not exist in table 'master.UnitMaster'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedUnitId;
                        }
                    }

                    // Special handling for ExhaustPressureUnitID and PressureUnitID in MechanicalDBO (resolve by unit name)
                    if (isMechanicalDBO &&
                        (string.Equals(mapping.SqlColumnName, "ExhaustPressureUnitID", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.SqlColumnName, "PressureUnitID", StringComparison.OrdinalIgnoreCase)) &&
                        value != DBNull.Value && value != null)
                    {
                        if (long.TryParse(value.ToString()?.Trim(), out var unitNumeric) && unitNumeric == 0)
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!unitIdCache.TryGetValue(valueKey, out var resolvedUnitId))
                            {
                                // Not in cache, resolve from database
                                resolvedUnitId = await ResolveUnitIdByNameAsync(
                                    connection,
                                    transaction,
                                    value,
                                    "master",
                                    "UnitMaster",
                                    "UnitName",
                                    cancellationToken);

                                // Cache the result (even if null)
                                unitIdCache[valueKey] = resolvedUnitId;
                            }

                            if (resolvedUnitId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: Unit '{value}' does not exist in table 'master.UnitMaster'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedUnitId;
                        }
                    }

                    // Special handling for CleanlinessFactor in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CleanlinessFactor", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCleanlinessFactorValue(value, mapping.IsNullable);
                    }

                    // Special handling for FoulingFactor in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "FoulingFactor", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOFoulingFactorValue(value, mapping.IsNullable);
                    }

                    // Special handling for PluggingMargin in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "PluggingMargin", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOPluggingMarginValue(value, mapping.IsNullable);
                    }

                    // Special handling for CWInletTemperature in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CWInletTemperature", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCWInletTemperatureValue(value, mapping.IsNullable);
                    }

                    // Special handling for CWOutletTemperature in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CWOutletTemperature", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCWOutletTemperatureValue(value, mapping.IsNullable);
                    }

                    // Special handling for CWSupplyPressure in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CWSupplyPressure", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCWSupplyPressureValue(value, mapping.IsNullable);
                    }

                    // Special handling for CWDesignPressure in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CWDesignPressure", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCWDesignPressureValue(value, mapping.IsNullable);
                    }

                    // Special handling for CWVelocity in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CWVelocity", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOCWVelocityValue(value, mapping.IsNullable);
                    }

                    // Special handling for VacuumBreakerValve in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "VacuumBreakerValve", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOVacuumBreakerValveValue(value, mapping.IsNullable);
                    }

                    // Special handling for Quantity in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Quantity", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOQuantityValue(value, mapping.IsNullable);
                    }

                    // Special handling for MaterialOfCasing in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "MaterialOfCasing", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOMaterialOfCasingValue(value, mapping.IsNullable);
                    }

                    // Special handling for AdditionalBOP in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "AdditionalBOP", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOAdditionalBOPValue(value, mapping.IsNullable);
                    }

                    // Special handling for RatedDifferentialHead in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "RatedDifferentialHead", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORatedDifferentialHeadValue(value, mapping.IsNullable);
                    }

                    // Special handling for FlowRating in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "FlowRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOFlowRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for InterAfterCondenser in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "InterAfterCondenser", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOInterAfterCondenserValue(value, mapping.IsNullable);
                    }

                    // Special handling for StartupEjector in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "StartupEjector", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOStartupEjectorValue(value, mapping.IsNullable);
                    }

                    // Special handling for MainEjector in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "MainEjector", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOMainEjectorValue(value, mapping.IsNullable);
                    }

                    // Special handling for EjectorNozzle in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "EjectorNozzle", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOEjectorNozzleValue(value, mapping.IsNullable);
                    }

                    // Special handling for TubesOfInterAfterCondenser in MechanicalDBO (uses same conversion as EjectorNozzle)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "TubesOfInterAfterCondenser", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOEjectorNozzleValue(value, mapping.IsNullable);
                    }

                    // Special handling for TubesSheetOfInterAfterCondenser in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "TubesSheetOfInterAfterCondenser", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOTubesSheetOfInterAfterCondenserValue(value, mapping.IsNullable);
                    }

                    // Special handling for ShellOfInterAfterCondenser in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "ShellOfInterAfterCondenser", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOShellOfInterAfterCondenserValue(value, mapping.IsNullable);
                    }

                    // Special handling for GlandSealing in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "GlandSealing", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOGlandSealingValue(value, mapping.IsNullable);
                    }

                    // Special handling for EjectionSystemDuringStartup in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "EjectionSystemDuringStartup", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOEjectionSystemDuringStartupValue(value, mapping.IsNullable);
                    }

                    // Special handling for WaterBoxes in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "WaterBoxes", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOWaterBoxesValue(value, mapping.IsNullable);
                    }

                    // Special handling for Shell in MechanicalDBO (uses same conversion as WaterBoxes)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Shell", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOWaterBoxesValue(value, mapping.IsNullable);
                    }

                    // Special handling for HotelWellRetentionTime in MechanicalDBO (uses same conversion as WaterBoxes)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "HotelWellRetentionTime", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOWaterBoxesValue(value, mapping.IsNullable);
                    }

                    // Special handling for Tubes in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Tubes", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOTubesValue(value, mapping.IsNullable);
                    }

                    // Special handling for GlandVentShell in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "GlandVentShell", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOGlandVentShellValue(value, mapping.IsNullable);
                    }

                    // Special handling for GlandVentTubes in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "GlandVentTubes", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOGlandVentTubesValue(value, mapping.IsNullable);
                    }

                    // Special handling for TubeSheets in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "TubeSheets", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOTubeSheetsValue(value, mapping.IsNullable);
                    }

                    // Special handling for Baffles in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Baffles", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOBafflesValue(value, mapping.IsNullable);
                    }

                    // Special handling for SafetyDeviceForCondenser in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "SafetyDeviceForCondenser", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOSafetyDeviceForCondenserValue(value, mapping.IsNullable);
                    }

                    // Special handling for Blower in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Blower", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOBlowerValue(value, mapping.IsNullable);
                    }

                    // Special handling for EjectionSystemForContinuous in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "EjectionSystemForContinuous", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOEjectionSystemForContinuousValue(value, mapping.IsNullable);
                    }

                    // Special handling for AutoGlandSealingSystem in MechanicalDBO
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "AutoGlandSealingSystem", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOAutoGlandSealingSystemValue(value, mapping.IsNullable);
                    }

                    // Special handling for GlandVentTubesSheet in MechanicalDBO (uses same conversion as SafetyDeviceForCondenser)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "GlandVentTubesSheet", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBOSafetyDeviceForCondenserValue(value, mapping.IsNullable);
                    }

                    // Special handling for ReliefValve in MechanicalDBO (uses RequiredNotRequired conversion)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "ReliefValve", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORequiredNotRequiredValue(value, mapping.IsNullable);
                    }

                    // Special handling for Rotometer in MechanicalDBO (uses RequiredNotRequired conversion)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "Rotometer", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORequiredNotRequiredValue(value, mapping.IsNullable);
                    }

                    // Special handling for CrossOverduct in MechanicalDBO (uses RequiredNotRequired conversion)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "CrossOverduct", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORequiredNotRequiredValue(value, mapping.IsNullable);
                    }

                    // Special handling for DumpProvision in MechanicalDBO (uses RequiredNotRequired conversion)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "DumpProvision", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORequiredNotRequiredValue(value, mapping.IsNullable);
                    }

                    // Special handling for LPGlandSealingAndDesuperheater in MechanicalDBO (uses RequiredNotRequired conversion)
                    if (isMechanicalDBO &&
                        string.Equals(mapping.SqlColumnName, "LPGlandSealingAndDesuperheater", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformMechanicalDBORequiredNotRequiredValue(value, mapping.IsNullable);
                    }

                    // Special handling for OrderTransmittalID column in OrderTransmittal child tables and ElectricalInstrumentationDBO tables (FK to OrderTransmittal)
                    // Handle Excel column k__ot_sel_ot_rec_bpp, ot_sel_ot_rec_bpp, or SQL column OrderTransmittalID
                    // Resolve by OrderTransmittalID (numeric) or by RecordNo (string)
                    if (((isOrderTransmittal && !string.Equals(tableName, "OrderTransmittal", StringComparison.OrdinalIgnoreCase)) ||
                         isOrderTransmittalNotes ||
                         isElectricalInstrumentationDBO) &&
                        (string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.ExcelColumnName, "k__ot_sel_ot_rec_bpp", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.ExcelColumnName, "ot_sel_ot_rec_bpp", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.ExcelColumnName, "record_id", StringComparison.OrdinalIgnoreCase)))
                    {
                        // If value is NULL or DBNull, keep it as NULL
                        if (value == DBNull.Value || value == null)
                        {
                            value = DBNull.Value;
                        }
                        // If value is numeric 0 or string "0", convert to NULL
                        else if ((value is int intVal && intVal == 0) ||
                                 (value is long longVal && longVal == 0) ||
                                 (value is short shortVal && shortVal == 0) ||
                                 (long.TryParse(value.ToString()?.Trim(), out var orderTransmittalNumeric) && orderTransmittalNumeric == 0))
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            if (isElectricalInstrumentationDBO &&
                                long.TryParse(value.ToString()?.Trim(), out var directOrderTransmittalId) &&
                                directOrderTransmittalId > 0)
                            {
                                value = directOrderTransmittalId;
                            }
                            else
                            {
                                var valueKey = value.ToString()?.Trim() ?? string.Empty;

                                // Check cache first
                                if (!orderTransmittalIdCache.TryGetValue(valueKey, out var resolvedOrderTransmittalId))
                                {
                                    // Not in cache, resolve from database
                                    resolvedOrderTransmittalId = await ResolveOrderTransmittalIdAsync(
                                        connection,
                                        transaction,
                                        value,
                                        cancellationToken);

                                    // Cache the result (even if null)
                                    orderTransmittalIdCache[valueKey] = resolvedOrderTransmittalId;
                                }

                                if (resolvedOrderTransmittalId == null)
                                {
                                    errorColumn = mapping.ExcelColumnName;
                                    errorValue = value;
                                    errorMessage = $"Foreign key constraint violation: OrderTransmittalID '{value}' does not exist in table 'bp.OrderTransmittal'";
                                    skipRow = true;
                                    break;
                                }

                                value = resolvedOrderTransmittalId;
                            }
                        }
                    }

                    // Special handling for ElectricalInstrumentationDBO columns
                    if (isElectricalInstrumentationDBO && value != DBNull.Value && value != null)
                    {
                        var colName = mapping.SqlColumnName;
                        if (string.Equals(colName, "Status", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformStatusValue(value, mapping.IsNullable);
                        }

                        if (string.Equals(colName, "IndividualCTPTForTVM", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformElectricalDBOCtPtValue(value, mapping.IsNullable);
                        }
                        //else if (colName.EndsWith("ScopeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformMechanicalDBOScopeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("AlternatorMakeID", StringComparison.OrdinalIgnoreCase) || colName.Equals("VMS_MakeID", StringComparison.OrdinalIgnoreCase) || colName.Equals("sync_MakeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOMakeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("StandardID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorStandardID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOStandardValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("VoltageAlternatorID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorVoltageID", StringComparison.OrdinalIgnoreCase) || colName.Equals("Battery_VoltageRatingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOVoltageValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("EnclosureID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorEnclosureID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOEnclosureValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("DesignTemperatureID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorDesignTempID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBODesignTempValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("RatedPFID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBORatedPfValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TemperatureRaiseID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorTempRiseID", StringComparison.OrdinalIgnoreCase) || colName.Equals("TransformerTempRiseID", StringComparison.OrdinalIgnoreCase) || colName.Equals("TemperatureRiseID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTempRiseValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("InsulationClassID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorInsulationClassID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOInsulationValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TerminalBoxToSuitID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorTBToSuitID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTerminalBoxValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.EndsWith("CertificationID", StringComparison.OrdinalIgnoreCase) || colName.EndsWith("CertID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOCertValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("NeutralCTStarFormationID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorNeutralCtStarID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBONeutralCtValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("PhaseSideCTLocationID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorPhaseSideCtID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOPhaseCtValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("ContinuousOverloadID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorOverloadID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOOverloadValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("NoiseLevelID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorNoiseLevelID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBONoiseValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("SlipRingID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorSlipRingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOSlipRingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("PMGID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorPMGID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOPMGValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TestsID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorTestsID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTestsValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("CoolingMethodID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorCoolingMethodID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOCoolingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("CoolerConfigurationID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorCoolerConfigID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOCoolerConfigValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("CoolerTubesMOCID", StringComparison.OrdinalIgnoreCase) || colName.Equals("AlternatorCoolerTubesMocID", StringComparison.OrdinalIgnoreCase) || colName.EndsWith("_MOCID", StringComparison.OrdinalIgnoreCase) || colName.EndsWith("MaterialID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //     if (colName.Contains("BusBar", StringComparison.OrdinalIgnoreCase))
                        //        value = TransformElectricalInstrumentationDBOBusBarMocValue(value, mapping.IsNullable);
                        //     else
                        //        value = TransformElectricalInstrumentationDBOMocValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Contains("IPRatingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOIPRatingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("ControlModeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOControlModeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Contains("RelayTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBORelayTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Contains("SyncTypeOfSynchronizerID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOSyncTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Contains("AccuracyID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOMeterAccuracyValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TVMTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTvmTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("SyncPQMID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOPqmValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TransformerTypeOfPanelID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTransformerTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Contains("FaultRatingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOFaultRatingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("SwitchGearBreakerTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBreakerTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("MotorControlConstTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOMccConstTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TCP_TypeOfControlPanelID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTcpTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TCP_RedundancyID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTcpRedundancyValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TGP_TypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTgpTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("VMS_VibrationMeasTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOVmsVibrValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("AVRStandbyExcitationID", StringComparison.OrdinalIgnoreCase) || 
                        //         colName.Equals("SyncPQMID", StringComparison.OrdinalIgnoreCase) ||
                        //         colName.EndsWith("RequiredID", StringComparison.OrdinalIgnoreCase) ||
                        //         colName.Equals("LASCPT_PartOfBreakerID", StringComparison.OrdinalIgnoreCase) ||
                        //         colName.Equals("PLCBasedInstrumentsID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOYessNoRequiredValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("Battery_CapacityID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBatteryCapacityValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TCP_SpecificationID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTcpSpecificationValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TCP_CommunicationTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTcpCommunicationTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TVMMountingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTvmMountingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("LTPowerCablingID", StringComparison.OrdinalIgnoreCase) ||
                        //         colName.Equals("ControlCablingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOLtPowerCablingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("HTPowerCablingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOHtPowerCablingValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("LTPowerCableMOCID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOLtPowerCableMocValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("BusDuctID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBusDuctValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("ControlCableMOCID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOControlCableMocValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("HTPowerCableMOCID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOHtPowerCableMocValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("BusDuctTypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBusDuctTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("Battery_TypeID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBatteryTypeValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("Battery_TypeOfChargerID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOBatteryTypeOfChargerValue(value, mapping.IsNullable);
                        //}
                        //else if (colName.Equals("TCP_SILRatingID", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    value = TransformElectricalInstrumentationDBOTcpSilRatingValue(value, mapping.IsNullable);
                        //}
                    }

                    // Special handling for OrderTransmittalRecordID column in BPComments table (FK to OrderTransmittal)
                    // Resolve by OrderTransmittalID (numeric) or by RecordNo (string)
                    if (isBPComments &&
                        string.Equals(mapping.SqlColumnName, "OrderTransmittalRecordID", StringComparison.OrdinalIgnoreCase))
                    {
                        // If value is NULL or DBNull, keep it as NULL
                        if (value == DBNull.Value || value == null)
                        {
                            value = DBNull.Value;
                        }
                        // If value is numeric 0 or string "0", convert to NULL
                        else if ((value is int intVal && intVal == 0) ||
                                 (value is long longVal && longVal == 0) ||
                                 (value is short shortVal && shortVal == 0) ||
                                 (long.TryParse(value.ToString()?.Trim(), out var orderTransmittalNumeric) && orderTransmittalNumeric == 0))
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!orderTransmittalIdCache.TryGetValue(valueKey, out var resolvedOrderTransmittalId))
                            {
                                // Not in cache, resolve from database
                                resolvedOrderTransmittalId = await ResolveOrderTransmittalIdAsync(
                                    connection,
                                    transaction,
                                    value,
                                    cancellationToken);

                                // Cache the result (even if null)
                                orderTransmittalIdCache[valueKey] = resolvedOrderTransmittalId;
                            }

                            if (resolvedOrderTransmittalId == null)
                            {
                                // For BPComments, if OrderTransmittalRecordID doesn't exist, set to NULL instead of skipping row
                                value = DBNull.Value;
                            }
                            else
                            {
                                value = resolvedOrderTransmittalId;
                            }
                        }
                    }

                    // Special handling for FK columns in Turbine table
                    // Handle common FKs: ProjectID, OrderTransmittalID, etc.
                    if (isTurbine)
                    {
                        // Handle StatusId column
                        if (string.Equals(mapping.SqlColumnName, "StatusId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineStatusValue(value, mapping.IsNullable);
                        }
                        // Handle TypeOfTurbineId column
                        else if (string.Equals(mapping.SqlColumnName, "TypeOfTurbineId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineTypeValue(value, mapping.IsNullable);
                        }
                        // Handle HMBDSubmittedId column
                        else if (string.Equals(mapping.SqlColumnName, "HMBDSubmittedId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineHMBDSubmittedValue(value, mapping.IsNullable);
                        }
                        // Handle GearBoxScope column
                        else if (string.Equals(mapping.SqlColumnName, "GearBoxScope", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineGearBoxScopeValue(value, mapping.IsNullable);
                        }
                        // Handle GearBoxTypeID column
                        else if (string.Equals(mapping.SqlColumnName, "GearBoxTypeID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineGearBoxTypeValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle EfficiencyId column
                        else if (string.Equals(mapping.SqlColumnName, "EfficiencyId", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineEfficiencyValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle LubeOilScopeId column
                        else if (string.Equals(mapping.SqlColumnName, "LubeOilScopeId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineLubeOilScopeValue(value, mapping.IsNullable);
                        }
                        // Handle MOPDriveId column
                        else if (string.Equals(mapping.SqlColumnName, "MOPDriveId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineMOPDriveValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle ControlOilFilterId column
                        else if (string.Equals(mapping.SqlColumnName, "ControlOilFilterId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineControlOilFilterValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle OilCoolerId column
                        else if (string.Equals(mapping.SqlColumnName, "OilCoolerId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineOilCoolerValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle SSTypeId column
                        else if (string.Equals(mapping.SqlColumnName, "SSTypeId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilPipingObj = excelRow.Table.Columns.Contains("uot_lube_oil_piping_pd") ? excelRow["uot_lube_oil_piping_pd"] : null;
                            value = TransformTurbineSSTypeValue(value, mapping.IsNullable, lubeOilPipingObj);
                        }
                        // Handle DrivenEquipmentId column
                        else if (string.Equals(mapping.SqlColumnName, "DrivenEquipmentId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineDrivenEquipmentValue(value, mapping.IsNullable);
                        }
                        // Handle GovernorScope column
                        else if (string.Equals(mapping.SqlColumnName, "GovernorScope", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineGovernorScopeValue(value, mapping.IsNullable);
                        }
                        // Handle ReductionID column
                        else if (string.Equals(mapping.SqlColumnName, "ReductionID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineReductionValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle GearBox_NoiseLevelID column
                        else if (string.Equals(mapping.SqlColumnName, "GearBox_NoiseLevelID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineGearBoxNoiseLevelValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle HighSpeedScopeId column
                        else if (string.Equals(mapping.SqlColumnName, "HighSpeedScopeId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineHighSpeedScopeValue(value, mapping.IsNullable);
                        }
                        // Handle SecondaryGBDrivenEqId column
                        else if (string.Equals(mapping.SqlColumnName, "SecondaryGBDrivenEqId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineSecondaryGBDrivenEqValue(value, mapping.IsNullable);
                        }
                        // Handle ManufacturingStandardID column
                        else if (string.Equals(mapping.SqlColumnName, "ManufacturingStandardID", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineManufacturingStandardValue(value, mapping.IsNullable);
                        }
                        // Handle InletOrientationId column
                        else if (string.Equals(mapping.SqlColumnName, "InletOrientationId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineInletOrientationValue(value, mapping.IsNullable);
                        }
                        // Handle RotationDirectionID column
                        else if (string.Equals(mapping.SqlColumnName, "RotationDirectionID", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineRotationDirectionValue(value, mapping.IsNullable);
                        }
                        // Handle FrameId column
                        else if (string.Equals(mapping.SqlColumnName, "FrameId", StringComparison.OrdinalIgnoreCase))
                        {
                            var turbineTypeObj = excelRow.Table.Columns.Contains("uot_turbine_pd") ? excelRow["uot_turbine_pd"] : null;
                            value = TransformTurbineFrameValue(value, mapping.IsNullable, turbineTypeObj);
                        }
                        // Handle ScopeId column
                        else if (string.Equals(mapping.SqlColumnName, "ScopeId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineScopeValue(value, mapping.IsNullable);
                        }
                        // Handle NoiseLevelID column
                        else if (string.Equals(mapping.SqlColumnName, "NoiseLevelID", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineNoiseLevelValue(value, mapping.IsNullable);
                        }
                        // Handle VendorListID column
                        else if (string.Equals(mapping.SqlColumnName, "VendorListID", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineVendorListValue(value, mapping.IsNullable);
                        }
                        // Handle ExhaustOrientationId column
                        else if (string.Equals(mapping.SqlColumnName, "ExhaustOrientationId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineExhaustOrientationValue(value, mapping.IsNullable);
                        }
                        // Handle Governor column
                        else if (string.Equals(mapping.SqlColumnName, "Governor", StringComparison.OrdinalIgnoreCase))
                        {
                            var governorScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope6_pd") ? excelRow["uot_ttlscope6_pd"] : null;
                            value = TransformTurbineGovernorValue(value, mapping.IsNullable, governorScopeObj);
                        }
                        // Handle QAPID column
                        else if (string.Equals(mapping.SqlColumnName, "QAPID", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineQAPValue(value, mapping.IsNullable);
                        }
                        // Handle BarringGearID column
                        else if (string.Equals(mapping.SqlColumnName, "BarringGearID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineBarringGearValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle SecondaryGearBoxID column
                        else if (string.Equals(mapping.SqlColumnName, "SecondaryGearBoxID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineSecondaryGearBoxValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle Type1Id column
                        else if (string.Equals(mapping.SqlColumnName, "Type1Id", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineCouplingTypeValue(value, mapping.IsNullable);
                        }
                        // Handle AcousticHoodId column
                        else if (string.Equals(mapping.SqlColumnName, "AcousticHoodId", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineAcousticHoodValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle PrimarySecondaryGBId column
                        else if (string.Equals(mapping.SqlColumnName, "PrimarySecondaryGBId", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbinePrimarySecondaryGBValue(value, mapping.IsNullable);
                        }
                        // Handle ShortCircuitFactorID column
                        else if (string.Equals(mapping.SqlColumnName, "ShortCircuitFactorID", StringComparison.OrdinalIgnoreCase))
                        {
                            var gearboxScopeObj = excelRow.Table.Columns.Contains("gearbox_scope") ? excelRow["gearbox_scope"] : null;
                            value = TransformTurbineShortCircuitFactorValue(value, mapping.IsNullable, gearboxScopeObj);
                        }
                        // Handle LubeTypeId column
                        else if (string.Equals(mapping.SqlColumnName, "LubeTypeId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineLubeOilTypeValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle VapourExtractorId column
                        else if (string.Equals(mapping.SqlColumnName, "VapourExtractorId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineVapourExtractorValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle TubeMOCId column
                        else if (string.Equals(mapping.SqlColumnName, "TubeMOCId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineTubeMOCValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle OilCentrifugeId column
                        else if (string.Equals(mapping.SqlColumnName, "OilCentrifugeId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineOilCentrifugeValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle LubeOilPipingId column
                        else if (string.Equals(mapping.SqlColumnName, "LubeOilPipingId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineLubeOilPipingValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle OverHeadTankId column
                        else if (string.Equals(mapping.SqlColumnName, "OverHeadTankId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineOverHeadTankValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle Type2Id column (low speed coupling type)
                        else if (string.Equals(mapping.SqlColumnName, "Type2Id", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineCouplingTypeLowSpeedValue(value, mapping.IsNullable);
                        }
                        // Handle OilHeatersId column
                        else if (string.Equals(mapping.SqlColumnName, "OilHeatersId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineOilHeatersValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle IfRequiredMOCId column
                        else if (string.Equals(mapping.SqlColumnName, "IfRequiredMOCId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineIfRequiredMOCValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle DirtyOilTankId column
                        else if (string.Equals(mapping.SqlColumnName, "DirtyOilTankId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineDirtyOilTankValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle TubeSheetsId column
                        else if (string.Equals(mapping.SqlColumnName, "TubeSheetsId", StringComparison.OrdinalIgnoreCase))
                        {
                            var oilCoolerObj = excelRow.Table.Columns.Contains("uot_oil_cooler_pd") ? excelRow["uot_oil_cooler_pd"] : null;
                            value = TransformTurbineTubeSheetsValue(value, mapping.IsNullable, oilCoolerObj);
                        }
                        // Handle FoulingFactorId column
                        else if (string.Equals(mapping.SqlColumnName, "FoulingFactorId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineFoulingFactorValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle OilFilterId column
                        else if (string.Equals(mapping.SqlColumnName, "OilFilterId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineOilFilterValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle AMOTTCVId column
                        else if (string.Equals(mapping.SqlColumnName, "AMOTTCVId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineAMOTTCVValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle PluggingMarginId column
                        else if (string.Equals(mapping.SqlColumnName, "PluggingMarginId", StringComparison.OrdinalIgnoreCase))
                        {
                            var tubeSheetsObj = excelRow.Table.Columns.Contains("uot_mech_tubesheet_pd") ? excelRow["uot_mech_tubesheet_pd"] : null;
                            value = TransformTurbinePluggingMarginValue(value, mapping.IsNullable, tubeSheetsObj);
                        }
                        // Handle IfRequiredCapacityId column
                        else if (string.Equals(mapping.SqlColumnName, "IfRequiredCapacityId", StringComparison.OrdinalIgnoreCase))
                        {
                            var lubeOilScopeObj = excelRow.Table.Columns.Contains("uot_ttlscope5_pd") ? excelRow["uot_ttlscope5_pd"] : null;
                            value = TransformTurbineIfRequiredCapacityValue(value, mapping.IsNullable, lubeOilScopeObj);
                        }
                        // Handle MaterialOfConstruction column
                        else if (string.Equals(mapping.SqlColumnName, "MaterialOfConstruction", StringComparison.OrdinalIgnoreCase))
                        {
                            value = TransformTurbineMaterialOfConstructionValue(value, mapping.IsNullable);
                        }
                        // Handle ProjectID FK
                        else if (string.Equals(mapping.SqlColumnName, "ProjectID", StringComparison.OrdinalIgnoreCase))
                        {
                            // If value is NULL or DBNull, keep it as NULL
                            if (value == DBNull.Value || value == null)
                            {
                                value = DBNull.Value;
                            }
                            // If value is numeric 0 or string "0", convert to NULL
                            else if ((value is int intVal && intVal == 0) ||
                                     (value is long longVal && longVal == 0) ||
                                     (value is short shortVal && shortVal == 0) ||
                                     (long.TryParse(value.ToString()?.Trim(), out var projectNumeric) && projectNumeric == 0))
                            {
                                value = DBNull.Value;
                            }
                            else
                            {
                                var valueKey = value.ToString()?.Trim() ?? string.Empty;

                                // Check cache first
                                if (!projectIdCache.TryGetValue(valueKey, out var resolvedProjectId))
                                {
                                    // Not in cache, resolve from database
                                    resolvedProjectId = await ResolveProjectIdAsync(
                                        connection,
                                        transaction,
                                        value,
                                        cancellationToken);

                                    // Cache the result (even if null)
                                    projectIdCache[valueKey] = resolvedProjectId;
                                }

                                if (resolvedProjectId == null)
                                {
                                    // For Turbine, if ProjectID doesn't exist, set to NULL instead of skipping row
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    value = resolvedProjectId;
                                }
                            }
                        }
                        // Handle OrderTransmittalID FK
                        else if (string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase))
                        {
                            // If value is NULL or DBNull, keep it as NULL
                            if (value == DBNull.Value || value == null)
                            {
                                value = DBNull.Value;
                            }
                            // If value is numeric 0 or string "0", convert to NULL
                            else if ((value is int intVal && intVal == 0) ||
                                     (value is long longVal && longVal == 0) ||
                                     (value is short shortVal && shortVal == 0) ||
                                     (long.TryParse(value.ToString()?.Trim(), out var orderTransmittalNumeric) && orderTransmittalNumeric == 0))
                            {
                                value = DBNull.Value;
                            }
                            else
                            {
                                var valueKey = value.ToString()?.Trim() ?? string.Empty;

                                // Check cache first
                                if (!orderTransmittalIdCache.TryGetValue(valueKey, out var resolvedOrderTransmittalId))
                                {
                                    // Not in cache, resolve from database
                                    resolvedOrderTransmittalId = await ResolveOrderTransmittalIdAsync(
                                        connection,
                                        transaction,
                                        value,
                                        cancellationToken);

                                    // Cache the result (even if null)
                                    orderTransmittalIdCache[valueKey] = resolvedOrderTransmittalId;
                                }

                                if (resolvedOrderTransmittalId == null)
                                {
                                    // For Turbine, if OrderTransmittalID doesn't exist, set to NULL instead of skipping row
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    value = resolvedOrderTransmittalId;
                                }
                            }
                        }
                        // Handle other FK columns ending with "ID" - generic FK lookup
                        // This will handle columns like CustomerID, VendorID, etc. if they exist
                        else if (mapping.SqlColumnName.EndsWith("ID", StringComparison.OrdinalIgnoreCase) &&
                                 !mapping.IsIdentity && // Exclude primary key columns
                                 !string.Equals(mapping.SqlColumnName, "ProjectID", StringComparison.OrdinalIgnoreCase) &&
                                 !string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase))
                        {
                            // If value is NULL or DBNull, keep it as NULL
                            if (value == DBNull.Value || value == null)
                            {
                                value = DBNull.Value;
                            }
                            // If value is numeric 0 or string "0", convert to NULL
                            else if ((value is int intVal && intVal == 0) ||
                                     (value is long longVal && longVal == 0) ||
                                     (value is short shortVal && shortVal == 0) ||
                                     (long.TryParse(value.ToString()?.Trim(), out var fkNumeric) && fkNumeric == 0))
                            {
                                value = DBNull.Value;
                            }
                            // For other FK columns, we'll let SQL Server validate them
                            // If they fail, the row will be skipped with an error
                        }
                    }






                    // Special handling for ContractClearance and LetterOfCorrespondence foreign keys - if they don't exist, set to NULL
                    if ((isContractClearance || isLetterOfCorrespondence) && 
                        value != DBNull.Value && value != null &&
                        !string.IsNullOrEmpty(mapping.ForeignKeyTableName))
                    {
                        var isValid = await ValidateForeignKeyValueAsync(
                            connection,
                            transaction,
                            mapping.ForeignKeyTableSchema ?? "dbo",
                            mapping.ForeignKeyTableName,
                            mapping.ForeignKeyColumnName,
                            value,
                            cancellationToken);

                        if (!isValid)
                        {
                            value = DBNull.Value;
                        }
                    }

                    // Special handling for OrderTransmittalID column in MechanicalDBO tables (FK to OrderTransmittal)
                    // Handle Excel column k__ot_sel_ot_rec_bpp or SQL column OrderTransmittalID
                    // Resolve by OrderTransmittalID (numeric) or by RecordNo (string)
                    if (isMechanicalDBO &&
                        (string.Equals(mapping.SqlColumnName, "OrderTransmittalID", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(mapping.ExcelColumnName, "k__ot_sel_ot_rec_bpp", StringComparison.OrdinalIgnoreCase)))
                    {
                        // If value is NULL or DBNull, keep it as NULL
                        if (value == DBNull.Value || value == null)
                        {
                            value = DBNull.Value;
                        }
                        // If value is numeric 0 or string "0", convert to NULL
                        else if ((value is int intVal && intVal == 0) ||
                                 (value is long longVal && longVal == 0) ||
                                 (value is short shortVal && shortVal == 0) ||
                                 (long.TryParse(value.ToString()?.Trim(), out var orderTransmittalNumeric) && orderTransmittalNumeric == 0))
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!orderTransmittalIdCache.TryGetValue(valueKey, out var resolvedOrderTransmittalId))
                            {
                                // Not in cache, resolve from database
                                resolvedOrderTransmittalId = await ResolveOrderTransmittalIdAsync(
                                    connection,
                                    transaction,
                                    value,
                                    cancellationToken);

                                // Cache the result (even if null)
                                orderTransmittalIdCache[valueKey] = resolvedOrderTransmittalId;
                            }

                            if (resolvedOrderTransmittalId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: OrderTransmittalID '{value}' does not exist in table 'bp.OrderTransmittal'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedOrderTransmittalId;
                        }
                    }

                    // Special handling for status column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for BGStatus column in OrderTransmittalLineItemBankGuarantee -> BankGuarantee migration
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "BGStatus", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalLineItemBankGuaranteeStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for status column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "Status", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalStatusValue(value, mapping.IsNullable);
                    }

                    // Special handling for OrderType column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "OrderType", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalOrderTypeValue(value, mapping.IsNullable);
                    }

                    // Special handling for Frequency column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "Frequency", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalFrequencyValue(value, mapping.IsNullable);
                    }

                    // Special handling for ServiceType column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ServiceType", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalServiceTypeValue(value, mapping.IsNullable);
                    }

                    // Special handling for INCOTerms column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "INCOTerms", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalINCOTermsValue(value, mapping.IsNullable);
                    }

                    // Special handling for ScopeOfSpares column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ScopeOfSpares", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalScopeOfSparesValue(value, mapping.IsNullable);
                    }

                    // Special handling for ScopeOfSeaworthyPacking column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ScopeOfSeaworthyPacking", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalScopeOfSeaworthyPackingValue(value, mapping.IsNullable);
                    }

                    // Special handling for SiteInsurance column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "SiteInsurance", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalSiteInsuranceValue(value, mapping.IsNullable);
                    }

                    // Special handling for MarineInsurance column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "MarineInsurance", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalMarineInsuranceValue(value, mapping.IsNullable);
                    }

                    // Special handling for TransitInsurance column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TransitInsurance", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTransitInsuranceValue(value, mapping.IsNullable);
                    }

                    // Special handling for ComprehensiveInsurance column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ComprehensiveInsurance", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalComprehensiveInsuranceValue(value, mapping.IsNullable);
                    }

                    // Special handling for StatutoryApproval column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "StatutoryApproval", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalStatutoryApprovalValue(value, mapping.IsNullable);
                    }

                    // Special handling for TransmittalTypeID column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TransmittalTypeID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTransmittalTypeIDValue(value, mapping.IsNullable);
                    }

                    // Special handling for TypesOfServicesEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TypesOfServicesEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTypesOfServicesEandCValue(value, mapping.IsNullable);
                    }

                    // Special handling for EotCraneFacilityEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "EotCraneFacilityEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalEotCraneFacilityEandCValue(value, mapping.IsNullable);
                    }

                    // Special handling for ErectionCraneEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ErectionCraneEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalErectionCraneEandCValue(value, mapping.IsNullable);
                    }

                    // Special handling for MobileCraneFacilityEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "MobileCraneFacilityEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for ConveyanceForEngineerEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ConveyanceForEngineerEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for UnloadingAtSiteEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "UnloadingAtSiteEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for GroutingEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "GroutingEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for GroutingMaterialSupplyEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "GroutingMaterialSupplyEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for StorageAtSiteEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "StorageAtSiteEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for ConstructionPowerWaterEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ConstructionPowerWaterEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for ErectionCableAndBaseEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ErectionCableAndBaseEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformByTTLOrCustomerValue(value, mapping.IsNullable);
                    }

                    // Special handling for TypeOfSparesEandC column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TypeOfSparesEandC", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTypeOfSparesEandCValue(value, mapping.IsNullable);
                    }

                    // Special handling for TypeOfWarranty column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TypeOfWarranty", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTypeOfWarrantyValue(value, mapping.IsNullable);
                    }

                    // Special handling for ReplacedPartsWarranty column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ReplacedPartsWarranty", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalReplacedPartsWarrantyValue(value, mapping.IsNullable);
                    }

                    // Special handling for EarthquakeZone column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "EarthquakeZone", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalEarthquakeZoneValue(value, mapping.IsNullable);
                    }

                    // Special handling for CoolingWater column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CoolingWater", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCoolingWaterValue(value, mapping.IsNullable);
                    }

                    // Special handling for MotorEfficiency column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "MotorEfficiency", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalMotorEfficiencyValue(value, mapping.IsNullable);
                    }

                    // Special handling for GeneratedVoltageRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "GeneratedVoltageRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalGeneratedVoltageRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for AuxiliaryVoltageRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "AuxiliaryVoltageRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalAuxiliaryVoltageRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for Environment column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "Environment", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalEnvironmentValue(value, mapping.IsNullable);
                    }

                    // Special handling for ScopeForCivil column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ScopeForCivil", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalScopeForCivilValue(value, mapping.IsNullable);
                    }

                    // Special handling for EPCorDirect column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "EPCorDirect", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalEPCorDirectValue(value, mapping.IsNullable);
                    }

                    // Special handling for TypeOfOrder column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TypeOfOrder", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTypeOfOrderValue(value, mapping.IsNullable);
                    }

                    // Special handling for CostOverrunRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CostOverrunRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCostOverrunRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for CostOverrunRiskRating column in OTRisk
                    // Mapping requested: R1 -> 0, R2 -> 1, R3 -> 2
                    if (isOTRisk &&
                        string.Equals(mapping.SqlColumnName, "CostOverrunRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOTRiskCostOverrunRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for ContractualDeliveryRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ContractualDeliveryRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalContractualDeliveryRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for CommercialTermsRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CommercialTermsRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCommercialTermsRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for CustomerRelationshipRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CustomerRelationshipRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCustomerRelationshipRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for FinancialHealthRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "FinancialHealthRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalFinancialHealthRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for AgreedPerformanceRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "AgreedPerformanceRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalAgreedPerformanceRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for WarrantyTermsRiskRating column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "WarrantyTermsRiskRating", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalWarrantyTermsRiskRatingValue(value, mapping.IsNullable);
                    }

                    // Special handling for CostOverrunImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CostOverrunImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCostOverrunImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for ContractualDeliveryImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ContractualDeliveryImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalContractualDeliveryImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for CommercialTermsImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CommercialTermsImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCommercialTermsImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for BusinessSector column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "BusinessSector", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalOthersBusinessSectorValue(value, mapping.IsNullable);
                    }

                    // Special handling for OthersBusinessSector column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "OthersBusinessSector", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalOthersBusinessSectorValue(value, mapping.IsNullable);
                    }

                    // Special handling for CustomerRelationshipImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "CustomerRelationshipImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCustomerRelationshipImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for FinancialHealthImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "FinancialHealthImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalFinancialHealthImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for AgreedPerformanceImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "AgreedPerformanceImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalAgreedPerformanceImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for WarrantyTermsImpact column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "WarrantyTermsImpact", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalWarrantyTermsImpactValue(value, mapping.IsNullable);
                    }

                    // Special handling for Currency column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "Currency", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalCurrencyValue(value, mapping.IsNullable);
                    }

                    // Special handling for Unit ID columns (resolve by unit name from UnitMaster)
                    if (isOrderTransmittal &&
                        UnitIdColumns.Contains(mapping.SqlColumnName) &&
                        value != DBNull.Value && value != null)
                    {
                        // If value is numeric 0, treat as NULL
                        if (long.TryParse(value.ToString()?.Trim(), out var unitNumeric) && unitNumeric == 0)
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!unitIdCache.TryGetValue(valueKey, out var resolvedUnitId))
                            {
                                // Not in cache, resolve from database
                                resolvedUnitId = await ResolveUnitIdByNameAsync(
                                    connection,
                                    transaction,
                                    value,
                                    "master",
                                    "UnitMaster",
                                    "UnitName",
                                    cancellationToken);

                                // Cache the result (even if null)
                                unitIdCache[valueKey] = resolvedUnitId;
                            }

                            if (resolvedUnitId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: Unit '{value}' does not exist in table 'master.UnitMaster'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedUnitId;
                        }
                    }

                    // Special handling for TaxesDutiesSpecify column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "TaxesDutiesSpecify", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalTaxesDutiesSpecifyValue(value, mapping.IsNullable);
                    }

                    // Special handling for ScopeOfFrieght column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ScopeOfFrieght", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalScopeOfFrieghtValue(value, mapping.IsNullable);
                    }

                    // Special handling for ScopeOfOptions column in OrderTransmittal
                    if (isOrderTransmittal &&
                        string.Equals(mapping.SqlColumnName, "ScopeOfOptions", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformOrderTransmittalScopeOfOptionsValue(value, mapping.IsNullable);
                    }

                    // Special handling for TypeOfGuarantee column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "TypeOfGuarantee", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeTypeOfGuaranteeValue(value, mapping.IsNullable);
                    }

                    // Special handling for WarrantyClause column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "WarrantyClause", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeWarrantyClauseValue(value, mapping.IsNullable);
                    }

                    // Special handling for GuaranteeAgainst column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "GuaranteeAgainst", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeGuaranteeAgainstValue(value, mapping.IsNullable);
                    }

                    // Special handling for DraftFormat column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "DraftFormat", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeDraftFormatValue(value, mapping.IsNullable);
                    }

                    // Special handling for BankGuaranteeType column in BankGuarantee
                    if (isBankGuarantee &&
                        string.Equals(mapping.SqlColumnName, "BankGuaranteeType", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        value = TransformBankGuaranteeTypeValue(value, mapping.IsNullable);
                    }

                    // Special handling for foreign key columns with int data type - directly convert Excel value to int
                    // Assume Excel already contains the ID value (not the name/description)
                    if (!string.IsNullOrWhiteSpace(mapping.ForeignKeyTableSchema) &&
                        !string.IsNullOrWhiteSpace(mapping.ForeignKeyTableName) &&
                        !string.IsNullOrWhiteSpace(mapping.ForeignKeyColumnName) &&
                        value != DBNull.Value && value != null)
                    {
                        // Check if the target column is int or bigint type
                        var isIntType = string.Equals(mapping.SqlDataType, "int", StringComparison.OrdinalIgnoreCase) ||
                                       string.Equals(mapping.SqlDataType, "bigint", StringComparison.OrdinalIgnoreCase) ||
                                       string.Equals(mapping.SqlDataType, "smallint", StringComparison.OrdinalIgnoreCase);

                        if (isIntType)
                        {
                            // Try to convert Excel value directly to int/bigint
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr))
                            {
                                if (long.TryParse(valueStr, out var longId))
                                {
                                    object? convertedValue = null;
                                    if (string.Equals(mapping.SqlDataType, "bigint", StringComparison.OrdinalIgnoreCase))
                                    {
                                        convertedValue = longId;
                                    }
                                    else if (string.Equals(mapping.SqlDataType, "int", StringComparison.OrdinalIgnoreCase))
                                    {
                                        convertedValue = (int)longId;
                                    }
                                    else if (string.Equals(mapping.SqlDataType, "smallint", StringComparison.OrdinalIgnoreCase))
                                    {
                                        convertedValue = (short)longId;
                                    }

                                    if (convertedValue != null)
                                    {
                                        // If value is 0, convert to NULL (0 often means "no value" for FK columns)
                                        if (IsZeroValue(convertedValue))
                                        {
                                            value = DBNull.Value;
                                        }
                                        else
                                        {
                                            // Validate that the FK value exists in the referenced table
                                            var fkExists = await ValidateForeignKeyValueAsync(
                                                connection,
                                                transaction,
                                                mapping.ForeignKeyTableSchema!,
                                                mapping.ForeignKeyTableName!,
                                                mapping.ForeignKeyColumnName!,
                                                convertedValue,
                                                cancellationToken);

                                            if (!fkExists)
                                            {
                                                if (isMinutesOfMeeting)
                                                {
                                                    value = DBNull.Value;
                                                }
                                                else
                                                {
                                                    errorColumn = mapping.ExcelColumnName;
                                                    errorValue = value;
                                                    errorMessage = $"Foreign key constraint violation: Value '{convertedValue}' does not exist in table '{mapping.ForeignKeyTableSchema}.{mapping.ForeignKeyTableName}.{mapping.ForeignKeyColumnName}'";
                                                    skipRow = true;
                                                    break;
                                                }
                                            }

                                            value = convertedValue;
                                        }
                                    }
                                }
                                // If parsing fails, let the normal conversion handle it (might throw error)
                            }
                        }
                    }
                    // Also validate known FK columns for OrderTransmittal table (even if FK metadata not populated)
                    else if (isOrderTransmittal && value != DBNull.Value && value != null)
                    {
                        // Check for CustomerContactID
                        if (string.Equals(mapping.SqlColumnName, "CustomerContactID", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var contactId))
                            {
                                // If value is 0, convert to NULL
                                if (contactId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerContacts",
                                        "CustomerContactID",
                                        contactId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: CustomerContactID '{contactId}' does not exist in table 'master.CustomerContacts'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                        // Check for CustomerContactID2
                        else if (string.Equals(mapping.SqlColumnName, "CustomerContactID2", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var contactId))
                            {
                                // If value is 0, convert to NULL
                                if (contactId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerContacts",
                                        "CustomerContactID",
                                        contactId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: CustomerContactID2 '{contactId}' does not exist in table 'master.CustomerContacts'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                        // Check for EndUserContactID
                        else if (string.Equals(mapping.SqlColumnName, "EndUserContactID", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var contactId))
                            {
                                // If value is 0, convert to NULL
                                if (contactId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerContacts",
                                        "CustomerContactID",
                                        contactId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: EndUserContactID '{contactId}' does not exist in table 'master.CustomerContacts'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                        // Check for EndUserContactID2
                        else if (string.Equals(mapping.SqlColumnName, "EndUserContactID2", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var contactId))
                            {
                                // If value is 0, convert to NULL
                                if (contactId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerContacts",
                                        "CustomerContactID",
                                        contactId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: EndUserContactID2 '{contactId}' does not exist in table 'master.CustomerContacts'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                        // Check for CustomerMasterID
                        else if (string.Equals(mapping.SqlColumnName, "CustomerMasterID", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var customerId))
                            {
                                // If value is 0, convert to NULL
                                if (customerId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerMaster",
                                        "CustomerID",
                                        customerId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: CustomerMasterID '{customerId}' does not exist in table 'master.CustomerMaster'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                        // Check for EndUserID
                        else if (string.Equals(mapping.SqlColumnName, "EndUserID", StringComparison.OrdinalIgnoreCase))
                        {
                            var valueStr = value.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(valueStr) && long.TryParse(valueStr, out var endUserId))
                            {
                                // If value is 0, convert to NULL
                                if (endUserId == 0)
                                {
                                    value = DBNull.Value;
                                }
                                else
                                {
                                    var fkExists = await ValidateForeignKeyValueAsync(
                                        connection,
                                        transaction,
                                        "master",
                                        "CustomerMaster",
                                        "CustomerID",
                                        endUserId,
                                        cancellationToken);

                                    if (!fkExists)
                                    {
                                        errorColumn = mapping.ExcelColumnName;
                                        errorValue = value;
                                        errorMessage = $"Foreign key constraint violation: EndUserID '{endUserId}' does not exist in table 'master.CustomerMaster'";
                                        skipRow = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    // Special handling for CustomerID column in CustomerContacts (FK to CustomerMaster)
                    // Resolve by CustomerID (numeric) or by CompanyCode/CompanyName (string)
                    if (isCustomerContacts &&
                        string.Equals(mapping.SqlColumnName, "CustomerID", StringComparison.OrdinalIgnoreCase) &&
                        value != DBNull.Value && value != null)
                    {
                        // If value is numeric 0, treat as NULL
                        if (long.TryParse(value.ToString()?.Trim(), out var customerNumeric) && customerNumeric == 0)
                        {
                            value = DBNull.Value;
                        }
                        else
                        {
                            var valueKey = value.ToString()?.Trim() ?? string.Empty;

                            // Check cache first
                            if (!customerIdCache.TryGetValue(valueKey, out var resolvedCustomerId))
                            {
                                // Not in cache, resolve from database
                                resolvedCustomerId = await ResolveCustomerIdByNameAsync(
                                    connection,
                                    transaction,
                                    value,
                                    cancellationToken);

                                // Cache the result (even if null)
                                customerIdCache[valueKey] = resolvedCustomerId;
                            }

                            if (resolvedCustomerId == null)
                            {
                                errorColumn = mapping.ExcelColumnName;
                                errorValue = value;
                                errorMessage = $"Foreign key constraint violation: Customer '{value}' does not exist in table 'master.CustomerMaster'";
                                skipRow = true;
                                break;
                            }

                            value = resolvedCustomerId;
                        }
                    }

                    if (value == DBNull.Value || value == null)
                    {
                        if (mapping.IsNullable)
                        {
                            newRow[mapping.SqlColumnName] = DBNull.Value;
                        }
                        else
                        {
                            // Skip rows with null values for non-nullable columns
                            errorColumn = mapping.ExcelColumnName;
                            errorValue = value;
                            errorMessage = $"Null value not allowed for non-nullable column '{mapping.SqlColumnName}'";
                            skipRow = true;
                            break;
                        }
                    }
                    else
                    {
                        // Convert value to match target column type
                        if (targetColumn != null)
                        {
                            newRow[mapping.SqlColumnName] = ConvertValue(value, targetColumn.DataType);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If conversion fails for a column, capture error details
                    errorColumn = mapping.ExcelColumnName;
                    try
                    {
                        errorValue = excelRow[mapping.ExcelColumnName];
                    }
                    catch
                    {
                        errorValue = "Unable to read value";
                    }
                    errorMessage = $"Type conversion error: {ex.Message}";
                    skipRow = true;
                    break;
                }
            }

            if (skipRow)
            {
                // Add row error detail
                rowErrors.Add(new Models.RowErrorDetail
                {
                    RowNumber = rowNumber,
                    ColumnName = errorColumn ?? "Unknown",
                    Value = errorValue,
                    ErrorMessage = errorMessage ?? "Unknown error",
                    RowData = rowData
                });
            }
            else
            {
                // Ensure IsDeleted is set to false (override if it was in Excel mappings)
                if (mappedTable.Columns.Contains("IsDeleted"))
                {
                    var isDeletedDataColumn = mappedTable.Columns["IsDeleted"];
                    if (isDeletedDataColumn != null && isDeletedDataColumn.DataType == typeof(bool))
                    {
                        newRow["IsDeleted"] = false;
                    }
                    else if (isDeletedDataColumn != null)
                    {
                        // For other types (int, bit as int, etc.), set to 0
                        newRow["IsDeleted"] = Convert.ChangeType(0, isDeletedDataColumn.DataType);
                    }
                }
                mappedTable.Rows.Add(newRow);
            }
        }

        return (mappedTable, rowErrors);
    }

    private Type GetNetTypeFromSqlType(string sqlDataType)
    {
        if (string.IsNullOrWhiteSpace(sqlDataType))
            return typeof(object);

        var type = sqlDataType.ToUpper().Trim();

        // Map SQL Server types to .NET types
        switch (type)
        {
            case "INT":
            case "INTEGER":
                return typeof(int);

            case "BIGINT":
                return typeof(long);

            case "SMALLINT":
                return typeof(short);

            case "TINYINT":
                return typeof(byte);

            case "BIT":
                return typeof(bool);

            case "DECIMAL":
            case "NUMERIC":
            case "MONEY":
            case "SMALLMONEY":
                return typeof(decimal);

            case "FLOAT":
            case "REAL":
                return typeof(double);

            case "DATE":
            case "DATETIME":
            case "DATETIME2":
            case "SMALLDATETIME":
                return typeof(DateTime);

            case "DATETIMEOFFSET":
                return typeof(DateTimeOffset);

            case "TIME":
                return typeof(TimeSpan);

            case "VARCHAR":
            case "NVARCHAR":
            case "CHAR":
            case "NCHAR":
            case "TEXT":
            case "NTEXT":
                return typeof(string);

            case "UNIQUEIDENTIFIER":
                return typeof(Guid);

            case "BINARY":
            case "VARBINARY":
            case "IMAGE":
                return typeof(byte[]);

            default:
                // For unknown types, return object to allow conversion
                return typeof(object);
        }
    }

    private async Task<string?> GetPrimaryKeyColumnNameAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        CancellationToken cancellationToken)
    {
        var query = @"
            SELECT TOP 1 ku.COLUMN_NAME
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                ON tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
                AND tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                AND tc.TABLE_SCHEMA = ku.TABLE_SCHEMA
                AND tc.TABLE_NAME = ku.TABLE_NAME
            WHERE tc.TABLE_SCHEMA = @SchemaName
                AND tc.TABLE_NAME = @TableName";

        try
        {
            await using var command = new SqlCommand(query, connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            command.Parameters.AddWithValue("@SchemaName", schemaName);
            command.Parameters.AddWithValue("@TableName", tableName);

            var result = await command.ExecuteScalarAsync(cancellationToken);
            if (result != null && result != DBNull.Value)
            {
                return result.ToString();
            }
        }
        catch
        {
            // If lookup fails, return null
        }

        return null;
    }

    private async Task<object?> ResolveUnitIdByNameAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        object excelValue,
        string unitTableSchema,
        string unitTableName,
        string unitNameColumnName,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return null;

        var valueStr = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(valueStr))
            return null;

        // Get the primary key column name dynamically
        var unitIdColumnName = await GetPrimaryKeyColumnNameAsync(
            connection,
            transaction,
            unitTableSchema,
            unitTableName,
            cancellationToken);

        if (string.IsNullOrWhiteSpace(unitIdColumnName))
        {
            // Fallback to common column names if primary key lookup fails
            unitIdColumnName = "UnitMasterID";
        }

        // If numeric, try to validate directly
        if (long.TryParse(valueStr, out var numericId))
        {
            var fkExists = await ValidateForeignKeyValueAsync(
                connection,
                transaction,
                unitTableSchema,
                unitTableName,
                unitIdColumnName,
                numericId,
                cancellationToken);

            return fkExists ? numericId : null;
        }

        // Otherwise, lookup by unit name (case-insensitive)
        var query = $@"
            SELECT TOP 1 [{unitIdColumnName}]
            FROM [{unitTableSchema}].[{unitTableName}]
            WHERE [{unitNameColumnName}] = @UnitName";

        await using var command = new SqlCommand(query, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@UnitName", valueStr);

        var result = await command.ExecuteScalarAsync(cancellationToken);
        if (result != null && result != DBNull.Value)
        {
            return result;
        }

        return null;
    }



    private async Task<object?> ResolveProjectTypeMasterIdByNameAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        object excelValue,
        string projectTypeTableSchema,
        string projectTypeTableName,
        string projectTypeNameColumnName,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return null;

        var valueStr = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(valueStr))
            return null;

        // Get the primary key column name dynamically
        var projectTypeMasterIdColumnName = await GetPrimaryKeyColumnNameAsync(
            connection,
            transaction,
            projectTypeTableSchema,
            projectTypeTableName,
            cancellationToken);

        if (string.IsNullOrWhiteSpace(projectTypeMasterIdColumnName))
        {
            // Fallback to common column names if primary key lookup fails
            projectTypeMasterIdColumnName = "ProjectTypeMasterID";
        }

        // If numeric, try to validate directly
        if (long.TryParse(valueStr, out var numericId))
        {
            var fkExists = await ValidateForeignKeyValueAsync(
                connection,
                transaction,
                projectTypeTableSchema,
                projectTypeTableName,
                projectTypeMasterIdColumnName,
                numericId,
                cancellationToken);

            return fkExists ? numericId : null;
        }

        // Otherwise, lookup by project type name (case-insensitive)
        var query = $@"
            SELECT TOP 1 [{projectTypeMasterIdColumnName}]
            FROM [{projectTypeTableSchema}].[{projectTypeTableName}]
            WHERE [{projectTypeNameColumnName}] = @ProjectTypeName";

        await using var command = new SqlCommand(query, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@ProjectTypeName", valueStr);

        var result = await command.ExecuteScalarAsync(cancellationToken);
        if (result != null && result != DBNull.Value)
        {
            return result;
        }

        return null;
    }

    private async Task<object?> ResolveCustomerIdByNameAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        object excelValue,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return null;

        var valueStr = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(valueStr))
            return null;

        // Get the primary key column name dynamically
        var customerIdColumnName = await GetPrimaryKeyColumnNameAsync(
            connection,
            transaction,
            "master",
            "CustomerMaster",
            cancellationToken);

        if (string.IsNullOrWhiteSpace(customerIdColumnName))
        {
            // Fallback to common column names if primary key lookup fails
            customerIdColumnName = "CustomerID";
        }

        // If numeric, try to validate directly
        if (long.TryParse(valueStr, out var numericId))
        {
            var fkExists = await ValidateForeignKeyValueAsync(
                connection,
                transaction,
                "master",
                "CustomerMaster",
                customerIdColumnName,
                numericId,
                cancellationToken);

            return fkExists ? numericId : null;
        }

        // Otherwise, lookup by CompanyCode first, then CompanyName (case-insensitive)
        // Try CompanyCode first
        var queryByCode = $@"
            SELECT TOP 1 [{customerIdColumnName}]
            FROM [master].[CustomerMaster]
            WHERE [CompanyCode] = @LookupValue";

        await using var commandByCode = new SqlCommand(queryByCode, connection, transaction);
        commandByCode.CommandTimeout = SqlCommandTimeout;
        commandByCode.Parameters.AddWithValue("@LookupValue", valueStr);

        var resultByCode = await commandByCode.ExecuteScalarAsync(cancellationToken);
        if (resultByCode != null && resultByCode != DBNull.Value)
        {
            return resultByCode;
        }

        // If not found by CompanyCode, try CompanyName
        var queryByName = $@"
            SELECT TOP 1 [{customerIdColumnName}]
            FROM [master].[CustomerMaster]
            WHERE [CompanyName] = @LookupValue";

        await using var commandByName = new SqlCommand(queryByName, connection, transaction);
        commandByName.CommandTimeout = SqlCommandTimeout;
        commandByName.Parameters.AddWithValue("@LookupValue", valueStr);

        var resultByName = await commandByName.ExecuteScalarAsync(cancellationToken);
        if (resultByName != null && resultByName != DBNull.Value)
        {
            return resultByName;
        }

        return null;
    }

    private async Task<object?> ResolveOrderTransmittalIdAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        object excelValue,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return null;

        var valueStr = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(valueStr))
            return null;

        // Get the primary key column name dynamically
        var orderTransmittalIdColumnName = await GetPrimaryKeyColumnNameAsync(
            connection,
            transaction,
            "bp",
            "OrderTransmittal",
            cancellationToken);

        if (string.IsNullOrWhiteSpace(orderTransmittalIdColumnName))
        {
            // Fallback to common column names if primary key lookup fails
            orderTransmittalIdColumnName = "OrderTransmittalID";
        }

        // If numeric, try to validate directly
        if (long.TryParse(valueStr, out var numericId))
        {
            var fkExists = await ValidateForeignKeyValueAsync(
                connection,
                transaction,
                "bp",
                "OrderTransmittal",
                orderTransmittalIdColumnName,
                numericId,
                cancellationToken);

            return fkExists ? numericId : null;
        }

        // Otherwise, lookup by RecordNo (case-insensitive)
        var query = $@"
            SELECT TOP 1 [{orderTransmittalIdColumnName}]
            FROM [bp].[OrderTransmittal]
            WHERE [RecordNo] = @RecordNo";

        await using var command = new SqlCommand(query, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@RecordNo", valueStr);

        var result = await command.ExecuteScalarAsync(cancellationToken);
        if (result != null && result != DBNull.Value)
        {
            return result;
        }

        return null;
    }

    private async Task<object?> ResolveProjectIdAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        object excelValue,
        CancellationToken cancellationToken)
    {
        if (excelValue == null || excelValue == DBNull.Value)
            return null;

        var valueStr = excelValue.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(valueStr))
            return null;

        // Get the primary key column name dynamically
        var projectIdColumnName = await GetPrimaryKeyColumnNameAsync(
            connection,
            transaction,
            "master",
            "Project",
            cancellationToken);

        if (string.IsNullOrWhiteSpace(projectIdColumnName))
        {
            // Fallback to common column names if primary key lookup fails
            projectIdColumnName = "ProjectID";
        }

        // If numeric, try to validate directly
        if (long.TryParse(valueStr, out var numericId))
        {
            var fkExists = await ValidateForeignKeyValueAsync(
                connection,
                transaction,
                "master",
                "Project",
                projectIdColumnName,
                numericId,
                cancellationToken);

            return fkExists ? numericId : null;
        }

        // Otherwise, lookup by ProjectName (case-insensitive)
        // Assume 'ProjectName' is the column if looking up by name
        var query = $@"
            SELECT TOP 1 [{projectIdColumnName}]
            FROM [master].[Project]
            WHERE [ProjectName] = @ProjectName";

        try
        {
            await using var command = new SqlCommand(query, connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            command.Parameters.AddWithValue("@ProjectName", valueStr);

            var result = await command.ExecuteScalarAsync(cancellationToken);
            if (result != null && result != DBNull.Value)
            {
                return result;
            }
        }
        catch
        {
            // If lookup fails (e.g. column doesn't exist or other error), return null implies project not found
        }

        return null;
    }

    private object TransformTurbineStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(statusStr, "Draft", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(statusStr, "Active", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(statusStr, "Inactive", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "Back Pressure", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Condensing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineHMBDSubmittedValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "Not Enclosed", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Enclosed", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Not Submitted", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineGearBoxScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "TTL", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Customer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Existing", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 4;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineGearBoxTypeValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Single Helical", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "Double Helical", StringComparison.OrdinalIgnoreCase)) return 2;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 1;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineEfficiencyValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Std", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "98.5%", StringComparison.OrdinalIgnoreCase)) return 2;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineLubeOilScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(valStr, "TTL", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Customer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Existing", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 4;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineMOPDriveValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Shaft Driven", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "AC Motor Driven", StringComparison.OrdinalIgnoreCase)) return 2;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineControlOilFilterValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Single", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "Duplex", StringComparison.OrdinalIgnoreCase)) return 2;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineOilCoolerValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Shell & Tube Type", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "Plate Type", StringComparison.OrdinalIgnoreCase)) return 2;
            else if (string.Equals(valStr, "Air Cooled", StringComparison.OrdinalIgnoreCase)) return 3;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineSSTypeValue(object value, bool isNullable, object? lubeOilPipingObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var pipingStr = lubeOilPipingObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(pipingStr, "Complete SS", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "SS 304", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "SS 316", StringComparison.OrdinalIgnoreCase)) return 2;
            else if (string.Equals(valStr, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        }
        else if (string.Equals(pipingStr, "Std", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineDrivenEquipmentValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "Alternator", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Compressor", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Fan", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Pump", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(valStr, "Others", StringComparison.OrdinalIgnoreCase)) return 5;

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineGovernorScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "TTL", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Customer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Existing", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 4;

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineReductionValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "Single", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "Double", StringComparison.OrdinalIgnoreCase)) return 2;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineGearBoxNoiseLevelValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;
        var scopeStr = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(scopeStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scopeStr, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
            else if (string.Equals(valStr, "90dBA", StringComparison.OrdinalIgnoreCase)) return 1;
            else if (string.Equals(valStr, "85dBA", StringComparison.OrdinalIgnoreCase)) return 2;
            else if (string.Equals(valStr, "109dBA(SPDP LT)", StringComparison.OrdinalIgnoreCase)) return 3;
            else if (string.Equals(valStr, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        }
        else if (string.Equals(scopeStr, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 0;
        }

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineHighSpeedScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "TTL", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Customer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Existing", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 4;

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineSecondaryGBDrivenEqValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var valStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(valStr, "Select", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(valStr, "TTL", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(valStr, "Customer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(valStr, "Existing", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(valStr, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 4;

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformTurbineManufacturingStandardValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return isNullable ? DBNull.Value : 0;

        return strValue.ToUpper() switch
        {
            "TTL" => 1,
            "API" => 2,
            "API-611" => 3,
            "API-612" => 4,
            "IEC" => 5,
            "OTHERS" => 6,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineInletOrientationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "STANDARD" => 1,
            "TOP" => 2,
            "BOTTOM" => 3,
            "SIDE" => 4,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineRotationDirectionValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "STANDARD" => 1,
            "CLOCK WISE" => 2,
            "COUNTER CLOCK WISE" => 3,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineFrameValue(object value, bool isNullable, object? turbineTypeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        var turbineType = turbineTypeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(turbineType, "Back Pressure", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "NOT APPLICABLE" => 1,
                "TST-1005-HB" => 2,
                "TST-1005-S" => 3,
                "TST-1015-HB" => 4,
                "TST-1015-L" => 5,
                "TST-1018-HB(SS)" => 6,
                "TST-1020-S" => 7,
                "TST-1025-HB" => 8,
                "TST-1030-EBM" => 9,
                "TST-1030-HB" => 10,
                "TST-1030-HBM" => 11,
                "TST-1030-M" => 12,
                "TST-1030-SBM" => 13,
                "TST-1050-EHB" => 14,
                "TST-1060-EB" => 15,
                "TST-1060-R" => 16,
                "TST-1060-REH" => 17,
                "TST-1060-SB" => 18,
                "TST-1100" => 19,
                "TST-1100-H" => 20,
                "TST-1150" => 21,
                "TST-1150-H" => 22,
                "TST-1150-HR" => 23,
                "TST-1150-RE" => 24,
                "TST-1220-EB" => 25,
                "TST-1250-HB" => 26,
                "TST-1250-SB" => 27,
                "TST-1300" => 28,
                "TST-1300-H" => 29,
                "TST-1300-HH" => 30,
                "TST-1300-LR" => 31,
                _ => 0
            };
        }
        else if (string.Equals(turbineType, "Condensing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "NOT APPLICABLE" => 1,
                "TST-2010" => 2,
                "TST-2025" => 3,
                "TST-2030" => 4,
                "TST-2060" => 5,
                "TST-2060-EI" => 6,
                "TST-2060-H" => 7,
                "TST-2060-R" => 8,
                "TST-2060-RE" => 9,
                "TST-2060-SR" => 10,
                "TST-2080" => 11,
                "TST-2080-AR" => 12,
                "TST-2080-G2-EI" => 13,
                "TST-2080-REH" => 14,
                "TST-2100" => 15,
                "TST-2100-G2" => 16,
                "TST-2100-G2-EI" => 17,
                "TST-2100-HL" => 18,
                "TST-2100-LR" => 19,
                "TST-2120" => 20,
                "TST-2120-H" => 21,
                "TST-2150" => 22,
                "TST-2150-H" => 23,
                "TST-2150-HH" => 24,
                "TST-2150-HHM" => 25,
                "TST-2150-IR" => 26,
                "TST-2160-ER" => 27,
                "TST-2160-HR" => 28,
                "TST-2180-G2" => 29,
                "TST-2180-G2-EI" => 30,
                "TST-2190-H" => 31,
                "TST-2200-H" => 32,
                "TST-2200-L" => 33,
                "TST-2210-H" => 34,
                "TST-2230-DE" => 35,
                "TST-2240-H" => 36,
                "TST-2250-HR" => 37,
                "TST-2280-HH" => 38,
                "TST-2300" => 39,
                "TST-2300-H" => 40,
                "TST-2300-HH" => 41,
                "TST-2300-HR" => 42,
                "TST-2300-LE" => 43,
                _ => 0
            };
        }
        else if (string.Equals(turbineType, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // mapping says: "Not Applicable": [{ id: 0, label: "Not Applicable", value: "Not Applicable" }]
        }

        return 0;
    }

    private object TransformTurbineScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "TTL" => 1,
            "CUSTOMER" => 2,
            "EXISTING" => 3,
            "NOT APPLICABLE" => 4,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineNoiseLevelValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "90DBA" => 1,
            "85DBA" => 2,
            "109DBA(SPDP LT)" => 3,
            "OTHERS" => 4,
            "NOT APPLICABLE" => 5,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineVendorListValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "TTL STANDARD" => 1,
            "TTL STANDARD + CUSTOMER/CONSULTANT" => 2,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineExhaustOrientationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "BOTTOM" => 1,
            "TOP" => 2,
            "AXIAL" => 3,
            "SIDE" => 4,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineGovernorValue(object value, bool isNullable, object? governorScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        var governorScope = governorScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(governorScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(governorScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "TTL STANDARD" => 1,
                "WOODWARD 505D" => 2,
                "WOODWARD 2301" => 3,
                "VOITH" => 4,
                "SITEC" => 5,
                "WOODWARD 505XT" => 6,
                "OTHERS" => 7,
                _ => 0
            };
        }
        else if (string.Equals(governorScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(governorScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "NOT APPLICABLE" => 1,
                _ => 0
            };
        }

        return 0;
    }

    private object TransformTurbineQAPValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "TTL STANDARD" => 1,
            "TTL STANDARD + CUSTOMER/CONSULTANT" => 2,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineBarringGearValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var gearboxScope = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(gearboxScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(gearboxScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "AUTO ENGAGE/AUTO DISENGAGE" => 1,
                "MANUAL ENGAGE/AUTO DISENGAGE" => 2,
                "NOT REQUIRED" => 3,
                _ => 0
            };
        }
        else if (string.Equals(gearboxScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(gearboxScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // mapping says: "Not Applicable": [{ id: 0, label: "Not Applicable", value: "Not Applicable" }]
        }

        return 0;
    }

    private object TransformTurbineSecondaryGearBoxValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var gearboxScope = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(gearboxScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(gearboxScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "REQUIRED" => 1,
                "NOT REQUIRED" => 2,
                _ => 0
            };
        }
        else if (string.Equals(gearboxScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(gearboxScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // mapping says: "Not Applicable": [{ id: 0, label: "Not Applicable", value: "Not Applicable" }]
        }

        return 0;
    }

    private object TransformTurbineCouplingTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "STANDARD" => 1,
            "FLEXIBLE MEMBRANE" => 2,
            "OTHERS" => 3,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineAcousticHoodValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var gearboxScope = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(gearboxScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(gearboxScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "GEAR BOX" => 1,
                "TURBINE & GEAR BOX" => 2,
                "TURBINE, GEARBOX & ALTERNATOR" => 3,
                "NOT REQUIRED" => 4,
                _ => 0
            };
        }
        else if (string.Equals(gearboxScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(gearboxScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // mapping says: "Not Applicable": [{ id: 0, label: "Not Applicable", value: "Not Applicable" }]
        }

        return 0;
    }

    private object TransformTurbinePrimarySecondaryGBValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "TTL" => 1,
            "CUSTOMER" => 2,
            "EXISTING" => 3,
            "NOT APPLICABLE" => 4,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineShortCircuitFactorValue(object value, bool isNullable, object? gearboxScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var gearboxScope = gearboxScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(gearboxScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(gearboxScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STD(6)" => 1,
                "OTHERS" => 2,
                _ => 0
            };
        }
        else if (string.Equals(gearboxScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(gearboxScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // mapping says: "Not Applicable": [{ id: 0, label: "Not Applicable", value: "Not Applicable" }]
        }

        return 0;
    }

    private object TransformTurbineLubeOilTypeValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STANDARD" => 1,
                "INTEGRAL WITH BASE PLATE" => 2,
                "SEPERATE OIL CONSOLE" => 3,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineVapourExtractorValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STD" => 1,
                "2*100% (2ND LOOSE SUPPLY)" => 2,
                "FRANKE TYPE OIL MIST ELIMINATOR" => 3,
                "AIR BREATHER" => 4,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineTubeMOCValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STANDARD" => 1,
                "OTHERS" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineOilCentrifugeValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "NOT IN TTL SCOPE" => 1,
                "STD (1000 LPH)" => 2,
                "OTHERS" => 3,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineLubeOilPipingValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STD" => 1,
                "COMPLETE SS" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineOverHeadTankValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "REQUIRED" => 1,
                "NOT REQUIRED" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineCouplingTypeLowSpeedValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        return strValue.ToUpper() switch
        {
            "RIGID" => 1,
            "FLEXIBLE" => 2,
            "GEAR" => 3,
            "STANDARD" => 4,
            "OTHERS" => 5,
            _ => isNullable ? DBNull.Value : 0
        };
    }

    private object TransformTurbineOilHeatersValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "REQUIRED" => 1,
                "NOT REQUIRED" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineIfRequiredMOCValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "CS" => 1,
                "CS WITH SS LINING" => 2,
                "COMPLETE SS" => 3,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineDirtyOilTankValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "REQUIRED" => 1,
                "NOT REQUIRED" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineTubeSheetsValue(object value, bool isNullable, object? oilCoolerObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        var oilCooler = oilCoolerObj?.ToString()?.Trim() ?? string.Empty;

        // For Shell & Tube Type or Air Cooled, apply the mapping
        if (string.Equals(oilCooler, "Shell & Tube Type", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(oilCooler, "Air Cooled", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STANDARD" => 1,
                "OTHERS" => 2,
                _ => 0
            };
        }
        // For Plate Type, return 0 (empty dependency map)
        else if (string.Equals(oilCooler, "Plate Type", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineFoulingFactorValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STANDARD" => 1,
                "OTHERS" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineOilFilterValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "STD ELEMENTS" => 1,
                "SS MESH" => 2,
                "OTHERS" => 3,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbineAMOTTCVValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "REQUIRED" => 1,
                "NOT REQUIRED" => 2,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    private object TransformTurbinePluggingMarginValue(object value, bool isNullable, object? tubeSheetsObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase))
            return 0;

        var tubeSheets = tubeSheetsObj?.ToString()?.Trim() ?? string.Empty;

        // For Standard or Others, apply the mapping
        if (string.Equals(tubeSheets, "Standard", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(tubeSheets, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "0 % (STD)" => 1,
                "5 %" => 2,
                "10 %" => 3,
                "OTHERS" => 4,
                _ => 0
            };
        }

        return 0;
    }

    private object TransformTurbineIfRequiredCapacityValue(object value, bool isNullable, object? lubeOilScopeObj)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(strValue) || strValue.Equals("Select", StringComparison.OrdinalIgnoreCase) || strValue.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;

        var lubeOilScope = lubeOilScopeObj?.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(lubeOilScope, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lubeOilScope, "Existing", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.ToUpper() switch
            {
                "1000 LTRS" => 1,
                "1500 LTRS" => 2,
                "2000 LTRS" => 3,
                "3000 LTRS" => 4,
                "OTHERS" => 5,
                _ => 0
            };
        }
        else if (string.Equals(lubeOilScope, "Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(lubeOilScope, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return 0;
    }

    


    private object TransformTurbineMaterialOfConstructionValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        var strValue = value.ToString()?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(strValue) || string.Equals(strValue, "Select", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (string.Equals(strValue, "TTL Standards", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(strValue, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        if (string.Equals(statusStr, "Active", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(statusStr, "Terminated", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformProjectStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        if (string.Equals(statusStr, "Approved", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(statusStr, "Created", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformProjectTemplateIdValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var templateStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        if (string.Equals(templateStr, "C Project Template -01", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel status string → SQL integer value
        if (string.Equals(statusStr, "Amended", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(statusStr, "Approved", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(statusStr, "Expired", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(statusStr, "Invoked", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(statusStr, "Notify_Concern", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(statusStr, "Pending", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(statusStr, "Send_for_auth_aprvl", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(statusStr, "send_for_bg_approval", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(statusStr, "Send_for_fin_acknow", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(statusStr, "Send_for__fin_rev", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else if (string.Equals(statusStr, "Sent_init_clarifi", StringComparison.OrdinalIgnoreCase))
        {
            return 11;
        }
        else if (string.Equals(statusStr, "Sent_Clarification", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }
        else if (string.Equals(statusStr, "Sent_for_Revision", StringComparison.OrdinalIgnoreCase))
        {
            return 13;
        }
        else if (string.Equals(statusStr, "Terminated", StringComparison.OrdinalIgnoreCase))
        {
            return 14;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalLineItemBankGuaranteeStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return isNullable ? DBNull.Value : 0;

        // If the value is already numeric (0/1), keep it.
        if (int.TryParse(value.ToString()?.Trim(), out var numericStatus))
        {
            if (numericStatus == 0 || numericStatus == 1)
                return numericStatus;
        }

        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(statusStr, "Issued", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(statusStr, "To Be Issued", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalStatusValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var statusStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel status string → SQL integer value
        if (string.Equals(statusStr, "Approved", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(statusStr, "CC_Pending", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(statusStr, "Convert_PROT_OT", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(statusStr, "Pending", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(statusStr, "Pending_Add_Info", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(statusStr, "PH_Apprval_Pending", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(statusStr, "PPC_in_P6_Pending", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(statusStr, "Proj_Head_Rev_Pend", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(statusStr, "Prop_Head_Rev_Pend", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(statusStr, "Sent_Clarification", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else if (string.Equals(statusStr, "Terminated", StringComparison.OrdinalIgnoreCase))
        {
            return 11;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalOrderTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var orderTypeStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel order type string → SQL integer value
        if (string.Equals(orderTypeStr, "Domestic", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(orderTypeStr, "Export", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(orderTypeStr, "Deemed Export", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(orderTypeStr, "Third Party Export", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 6;
        }
    }

    private object TransformOrderTransmittalFrequencyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var frequencyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel frequency string → SQL integer value
        if (string.Equals(frequencyStr, "50 Hz", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(frequencyStr, "60 Hz", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalServiceTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var serviceTypeStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel service type string → SQL integer value
        if (string.Equals(serviceTypeStr, "Turnkey", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(serviceTypeStr, "Supervision", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(serviceTypeStr, "Third party supervision", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 6;
        }
    }

    private object TransformOrderTransmittalINCOTermsValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var incotermsStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel INCOTerms string → SQL integer value
        if (string.Equals(incotermsStr, "Ex-works", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(incotermsStr, "FCA", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(incotermsStr, "CPT", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(incotermsStr, "CIP", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(incotermsStr, "DAP", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(incotermsStr, "DPU", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(incotermsStr, "DDP", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(incotermsStr, "FAS", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(incotermsStr, "FOB", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(incotermsStr, "CFR", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(incotermsStr, "CIF", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else if (string.Equals(incotermsStr, "FOR", StringComparison.OrdinalIgnoreCase))
        {
            return 11;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalScopeOfSparesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var scopeOfSparesStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel scope of spares string → SQL integer value
        if (string.Equals(scopeOfSparesStr, "Included in Order Value", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(scopeOfSparesStr, "Not in Scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(scopeOfSparesStr, "Separate Price", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalScopeOfSeaworthyPackingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var scopeOfSeaworthyPackingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel scope of seaworthy packing string → SQL integer value
        if (string.Equals(scopeOfSeaworthyPackingStr, "Included in the order value", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(scopeOfSeaworthyPackingStr, "Not in Scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(scopeOfSeaworthyPackingStr, "Separate Price", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalMarineInsuranceValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var marineInsuranceStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel marine insurance string → SQL integer value
        if (string.Equals(marineInsuranceStr, "TTL scope", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(marineInsuranceStr, "Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(marineInsuranceStr, "Not applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalSiteInsuranceValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var siteInsuranceStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel site insurance string → SQL integer value
        if (string.Equals(siteInsuranceStr, "TTL scope", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(siteInsuranceStr, "Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(siteInsuranceStr, "Not applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalTransitInsuranceValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var transitInsuranceStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel transit insurance string → SQL integer value
        if (string.Equals(transitInsuranceStr, "TTL scope", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(transitInsuranceStr, "Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(transitInsuranceStr, "Not applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalComprehensiveInsuranceValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var comprehensiveInsuranceStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel comprehensive insurance string → SQL integer value
        if (string.Equals(comprehensiveInsuranceStr, "TTL scope", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(comprehensiveInsuranceStr, "Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(comprehensiveInsuranceStr, "Not applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalStatutoryApprovalValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var statutoryApprovalStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel statutory approval string → SQL integer value
        if (string.Equals(statutoryApprovalStr, "TTL scope", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(statutoryApprovalStr, "Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(statutoryApprovalStr, "Not applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalTransmittalTypeIDValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var transmittalTypeIDStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel transmittal type string → SQL integer value
        if (string.Equals(transmittalTypeIDStr, "Order Transmittal", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(transmittalTypeIDStr, "Provisional Order Transmittal", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalTypesOfServicesEandCValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typesOfServicesEandCStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel types of services E&C string → SQL integer value
        if (string.Equals(typesOfServicesEandCStr, "Only supervision of E & C", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(typesOfServicesEandCStr, "Erection & Commissioning", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typesOfServicesEandCStr, "Third party supervision", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalTypeOfSparesEandCValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typeOfSparesEandCStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel type of spares E&C string → SQL integer value
        if (string.Equals(typeOfSparesEandCStr, "Commissioning (Standard)", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(typeOfSparesEandCStr, "2 Years Spares", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typeOfSparesEandCStr, "Additional Spares", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(typeOfSparesEandCStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalTypeOfWarrantyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typeOfWarrantyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel type of warranty string → SQL integer value
        if (string.Equals(typeOfWarrantyStr, "2 Crushing seasons", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(typeOfWarrantyStr, "12 months from the date of commissioning / 18 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typeOfWarrantyStr, "12 months from the date of commissioning / 24 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(typeOfWarrantyStr, "18 months from the date of commissioning / 24 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(typeOfWarrantyStr, "18 months from the date of commissioning / 36 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(typeOfWarrantyStr, "18 months from the date of commissioning / 42 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(typeOfWarrantyStr, "24 months from the date of commissioning / 30 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(typeOfWarrantyStr, "24 months from the date of commissioning / 36 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(typeOfWarrantyStr, "30 months from the date of commissioning / 36 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(typeOfWarrantyStr, "36 months from the date of commissioning / 42 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(typeOfWarrantyStr, "36 months from the date of commissioning / 60 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else if (string.Equals(typeOfWarrantyStr, "42 months from the date of commissioning / 48 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 11;
        }
        else if (string.Equals(typeOfWarrantyStr, "42 months from the date of commissioning / 60 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }
        else if (string.Equals(typeOfWarrantyStr, "48 months from the date of commissioning / 54 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 13;
        }
        else if (string.Equals(typeOfWarrantyStr, "54 months from the date of commissioning / 60 months from the date of dispatch (whichever is earlier)", StringComparison.OrdinalIgnoreCase))
        {
            return 14;
        }
        else if (string.Equals(typeOfWarrantyStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 15;
        }
        else if (string.Equals(typeOfWarrantyStr, "Under Warranty", StringComparison.OrdinalIgnoreCase))
        {
            return 16;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalReplacedPartsWarrantyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var replacedPartsWarrantyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel replaced parts warranty string → SQL integer value
        if (string.Equals(replacedPartsWarrantyStr, "Original Warranty", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(replacedPartsWarrantyStr, "12 Months from Replacement", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalEarthquakeZoneValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var earthquakeZoneStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel earthquake zone string → SQL integer value
        if (string.Equals(earthquakeZoneStr, "Safe Zone", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(earthquakeZoneStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalCoolingWaterValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var coolingWaterStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel cooling water string → SQL integer value
        if (string.Equals(coolingWaterStr, "Normal", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(coolingWaterStr, "Treated", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(coolingWaterStr, "Industrial", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(coolingWaterStr, "Sea Water", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalMotorEfficiencyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var motorEfficiencyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel motor efficiency string → SQL integer value
        if (string.Equals(motorEfficiencyStr, "IE2 (Std)", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(motorEfficiencyStr, "IE3", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(motorEfficiencyStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 20;
        }
    }

    private object TransformOrderTransmittalGeneratedVoltageRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(str, "380 V", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "400 V", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "415 V", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "660 V", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "3.3 KV", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "6.6 KV", StringComparison.OrdinalIgnoreCase))
            return 5;
        else if (string.Equals(str, "11 KV", StringComparison.OrdinalIgnoreCase))
            return 6;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 7;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalAuxiliaryVoltageRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(str, "380 V", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "400 V", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "415 V", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "660 V", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "3.3 KV", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "6.6 KV", StringComparison.OrdinalIgnoreCase))
            return 5;
        else if (string.Equals(str, "11 KV", StringComparison.OrdinalIgnoreCase))
            return 6;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 7;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalEnvironmentValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(str, "Dusty", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Acidic", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Other", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalScopeForCivilValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(str, "Yes", StringComparison.OrdinalIgnoreCase) || str == "1")
            return 1;
        else if (string.Equals(str, "No", StringComparison.OrdinalIgnoreCase) || str == "0")
            return 0;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalEPCorDirectValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var epcOrDirectStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel EPC or Direct string → SQL integer value
        if (string.Equals(epcOrDirectStr, "EPC", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(epcOrDirectStr, "Direct", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalTypeOfOrderValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typeOfOrderStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel type of order string → SQL integer value
        if (string.Equals(typeOfOrderStr, "Contract", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(typeOfOrderStr, "Agreement", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typeOfOrderStr, "LOI", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(typeOfOrderStr, "Purchase Order", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCostOverrunRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var costOverrunRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel cost overrun risk rating string → SQL integer value
        if (string.Equals(costOverrunRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(costOverrunRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(costOverrunRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOTRiskCostOverrunRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var ratingStr = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(ratingStr, "R1", StringComparison.OrdinalIgnoreCase))
            return 0;
        if (string.Equals(ratingStr, "R2", StringComparison.OrdinalIgnoreCase))
            return 1;
        if (string.Equals(ratingStr, "R3", StringComparison.OrdinalIgnoreCase))
            return 2;

        // Keep already-numeric values intact when Excel already stores 0/1/2.
        if (int.TryParse(ratingStr, out var numeric))
            return numeric;

        return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalCustomerRelationshipRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var customerRelationshipRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel customer relationship risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(customerRelationshipRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(customerRelationshipRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(customerRelationshipRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalFinancialHealthRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var financialHealthRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel financial health risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(financialHealthRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(financialHealthRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(financialHealthRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalAgreedPerformanceRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var agreedPerformanceRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel agreed performance risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(agreedPerformanceRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(agreedPerformanceRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(agreedPerformanceRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalWarrantyTermsRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var warrantyTermsRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel warranty terms risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(warrantyTermsRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(warrantyTermsRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(warrantyTermsRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCommercialTermsRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var commercialTermsRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel commercial terms risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(commercialTermsRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(commercialTermsRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(commercialTermsRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalContractualDeliveryRiskRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var contractualDeliveryRiskRatingStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel contractual delivery risk rating string → SQL integer value
        // Same mapping as CostOverrunRiskRating: R1→0, R2→1, R3→2
        if (string.Equals(contractualDeliveryRiskRatingStr, "R1", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(contractualDeliveryRiskRatingStr, "R2", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(contractualDeliveryRiskRatingStr, "R3", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCostOverrunImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var costOverrunImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel cost overrun impact string → SQL integer value
        if (string.Equals(costOverrunImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(costOverrunImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(costOverrunImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCommercialTermsImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var commercialTermsImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel commercial terms impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(commercialTermsImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(commercialTermsImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(commercialTermsImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalOthersBusinessSectorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var othersBusinessSectorStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel others business sector string → SQL integer value
        if (string.Equals(othersBusinessSectorStr, "Sugar", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(othersBusinessSectorStr, "Palm Oil", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(othersBusinessSectorStr, "Biomass", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(othersBusinessSectorStr, "Distillery", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(othersBusinessSectorStr, "Oil & Gas", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(othersBusinessSectorStr, "Cement", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(othersBusinessSectorStr, "Pulp & Paper", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(othersBusinessSectorStr, "Textile", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(othersBusinessSectorStr, "Waste to Energy", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(othersBusinessSectorStr, "Food & Beverage", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(othersBusinessSectorStr, "Chemical & Fertilizers", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else if (string.Equals(othersBusinessSectorStr, "Steel", StringComparison.OrdinalIgnoreCase))
        {
            return 11;
        }
        else if (string.Equals(othersBusinessSectorStr, "IPP", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }
        else if (string.Equals(othersBusinessSectorStr, "Carbon Black", StringComparison.OrdinalIgnoreCase))
        {
            return 13;
        }
        else if (string.Equals(othersBusinessSectorStr, "District Heating", StringComparison.OrdinalIgnoreCase))
        {
            return 14;
        }
        else if (string.Equals(othersBusinessSectorStr, "Pharmaceutical", StringComparison.OrdinalIgnoreCase))
        {
            return 15;
        }
        else if (string.Equals(othersBusinessSectorStr, "CHP", StringComparison.OrdinalIgnoreCase))
        {
            return 16;
        }
        else if (string.Equals(othersBusinessSectorStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 17;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCustomerRelationshipImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var customerRelationshipImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel customer relationship impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(customerRelationshipImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(customerRelationshipImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(customerRelationshipImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalFinancialHealthImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var financialHealthImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel financial health impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(financialHealthImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(financialHealthImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(financialHealthImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalAgreedPerformanceImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var agreedPerformanceImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel agreed performance impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(agreedPerformanceImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(agreedPerformanceImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(agreedPerformanceImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalWarrantyTermsImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var warrantyTermsImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel warranty terms impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(warrantyTermsImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(warrantyTermsImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(warrantyTermsImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalContractualDeliveryImpactValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var contractualDeliveryImpactStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel contractual delivery impact string → SQL integer value
        // Same mapping as CostOverrunImpact: Low→0, Medium→1, High→2
        if (string.Equals(contractualDeliveryImpactStr, "Low", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(contractualDeliveryImpactStr, "Medium", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(contractualDeliveryImpactStr, "High", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalCurrencyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;

        if (string.Equals(str, "INR", StringComparison.OrdinalIgnoreCase) || str == "₹")
            return 0;
        else if (string.Equals(str, "USD", StringComparison.OrdinalIgnoreCase) || str == "$")
            return 1;
        else if (string.Equals(str, "EUR", StringComparison.OrdinalIgnoreCase) || str == "€")
            return 2;
        else if (string.Equals(str, "GBP", StringComparison.OrdinalIgnoreCase) || str == "£")
            return 3;
        else if (string.Equals(str, "AUD", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "NZD", StringComparison.OrdinalIgnoreCase))
            return 5;
        else if (string.Equals(str, "CAD", StringComparison.OrdinalIgnoreCase))
            return 6;
        else if (string.Equals(str, "CHF", StringComparison.OrdinalIgnoreCase))
            return 7;
        else if (string.Equals(str, "JPY", StringComparison.OrdinalIgnoreCase))
            return 8;
        else if (string.Equals(str, "ZAR", StringComparison.OrdinalIgnoreCase))
            return 9;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformOrderTransmittalEotCraneFacilityEandCValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var eotCraneFacilityEandCStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel EOT crane facility E&C string → SQL integer value
        if (string.Equals(eotCraneFacilityEandCStr, "Available", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(eotCraneFacilityEandCStr, "Not Available", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 10;
        }
    }

    private object TransformOrderTransmittalErectionCraneEandCValue(object value, bool isNullable)
    {
        return TransformByTTLOrCustomerValue(value, isNullable);
    }

    // Common transformation method for columns with "By TTL"/"TTL" → 0 and "By Customer"/"Customer" → 1 mapping
    private object TransformByTTLOrCustomerValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var valueStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: "By TTL", "TTL", "By Triveni" → 0
        //          "By Customer", "Customer" → 1
        if (string.Equals(valueStr, "By TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(valueStr, "TTL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(valueStr, "By Triveni", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(valueStr, "By Customer", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(valueStr, "Customer", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalTaxesDutiesSpecifyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var taxesDutiesSpecifyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel taxes duties specify string → SQL integer value
        if (string.Equals(taxesDutiesSpecifyStr, "Included in the PO value", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(taxesDutiesSpecifyStr, "Extra as per Actual", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalScopeOfFrieghtValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var scopeOfFrieghtStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel scope of freight string → SQL integer value
        if (string.Equals(scopeOfFrieghtStr, "Included in the order value", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(scopeOfFrieghtStr, "In Purchaser scope", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(scopeOfFrieghtStr, "To be arranged by TTL on \"To Pay\" basis", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(scopeOfFrieghtStr, "To be arranged by TTL on 'To Pay' basis", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(scopeOfFrieghtStr, "Separate Price", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformOrderTransmittalScopeOfOptionsValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var scopeOfOptionsStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel scope of options string → SQL integer value
        if (string.Equals(scopeOfOptionsStr, "Included in the PO value", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        else if (string.Equals(scopeOfOptionsStr, "Extra as per Actual", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeTypeOfGuaranteeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typeStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel TypeOfGuarantee string → SQL integer value
        if (string.Equals(typeStr, "Advance Bank Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typeStr, "Perfomance Bank Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(typeStr, "Corporate Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(typeStr, "Corporate Performance Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(typeStr, "Counter Bank Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(typeStr, "Financial Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(typeStr, "Foreign Bank Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(typeStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(typeStr, "Corporate Bank Guarantee", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeWarrantyClauseValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var warrantyStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel WarrantyClause string → SQL integer value
        if (string.Equals(warrantyStr, "12/18 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(warrantyStr, "12/24 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(warrantyStr, "18/24 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(warrantyStr, "18/36 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(warrantyStr, "18/42 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else if (string.Equals(warrantyStr, "24/36 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }
        else if (string.Equals(warrantyStr, "36/60 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }
        else if (string.Equals(warrantyStr, "2 Crushing Season", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }
        else if (string.Equals(warrantyStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 9;
        }
        else if (string.Equals(warrantyStr, "12/36 months from the date of Commissioning or Dispatch whichever is earlier", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeGuaranteeAgainstValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var guaranteeStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel GuaranteeAgainst string → SQL integer value
        if (string.Equals(guaranteeStr, "Contract", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(guaranteeStr, "E&C", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(guaranteeStr, "Performance", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(guaranteeStr, "Supply", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else if (string.Equals(guaranteeStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeDraftFormatValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var draftFormatStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel DraftFormat string → SQL integer value
        if (string.Equals(draftFormatStr, "Customer", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(draftFormatStr, "Not Applicable", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else if (string.Equals(draftFormatStr, "Others", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }
        else if (string.Equals(draftFormatStr, "TTL", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    private object TransformBankGuaranteeTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // Convert to string and trim
        var typeStr = value.ToString()?.Trim() ?? string.Empty;

        // Case-insensitive comparison and transform
        // Mapping: Excel BankGuaranteeType string → SQL integer value
        if (string.Equals(typeStr, "New", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        else if (string.Equals(typeStr, "Amendment", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        else
        {
            // Default to NULL if column is nullable, otherwise 0
            return isNullable ? DBNull.Value : 0;
        }
    }

    // MechanicalDBO Transformation Methods
    private object TransformMechanicalDBOScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Water Cooled", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Air Cooled", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOPressureUnitValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "kg/cm²", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "kg/cm²g", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "barg", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "bara", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "ata", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "kPa", StringComparison.OrdinalIgnoreCase))
            return 5;
        else if (string.Equals(str, "MPa", StringComparison.OrdinalIgnoreCase))
            return 6;
        else if (string.Equals(str, "PSI", StringComparison.OrdinalIgnoreCase))
            return 7;
        else if (string.Equals(str, "kg/cm²a", StringComparison.OrdinalIgnoreCase))
            return 8;
        else if (string.Equals(str, "kg/m²", StringComparison.OrdinalIgnoreCase))
            return 9;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCleanlinessFactorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.85(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOFoulingFactorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std(0.00015)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOPluggingMarginValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0% std", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCWInletTemperatureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "32", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "33", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "34", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "35", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 5;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCWOutletTemperatureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "40", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "41", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "42", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 4;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCWSupplyPressureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std(3 Ata)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCWDesignPressureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std(6 Ata)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOCWVelocityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SS34(2.13 m/s)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Adm Brass(1.8 m/s)", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOVacuumBreakerValveValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL Scope", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Not Required", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOQuantityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "2 (std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "3", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOMaterialOfCasingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std - Cast Iron", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "CS", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOAdditionalBOPValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Cooling Tower", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Air Compressor", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "EOT", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Fire fighting system", StringComparison.OrdinalIgnoreCase))
            return 3;
        else if (string.Equals(str, "Deaerator", StringComparison.OrdinalIgnoreCase))
            return 4;
        else if (string.Equals(str, "SWAS", StringComparison.OrdinalIgnoreCase))
            return 5;
        else if (string.Equals(str, "BFWP", StringComparison.OrdinalIgnoreCase))
            return 6;
        else if (string.Equals(str, "Pumps", StringComparison.OrdinalIgnoreCase))
            return 7;
        else if (string.Equals(str, "Grouting Cement", StringComparison.OrdinalIgnoreCase))
            return 8;
        else if (string.Equals(str, "HVAC", StringComparison.OrdinalIgnoreCase))
            return 9;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 10;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBORatedDifferentialHeadValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std(80 mtrs)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOFlowRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std - 1.1 times condensor flow", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOInterAfterCondenserValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1 x 100% (Std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "2 x 100%", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOStartupEjectorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std 1 x 100%", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOMainEjectorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std 1 x 100%", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOEjectorNozzleValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SA 479 TP 304 (std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOTubesSheetOfInterAfterCondenserValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IS 2002(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "SA 516 Gr.70", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOShellOfInterAfterCondenserValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SA 106 Gr.B(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOGlandSealingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Auxillary steam line (through PRDS)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 1;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOEjectionSystemDuringStartupValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Auxillary Steams", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOWaterBoxesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IS2062(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOTubesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SS 304 ERW (std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOGlandVentShellValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SA106Gr.B(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOGlandVentTubesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SA106Gr.B(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOTubeSheetsValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Standard", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOBafflesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IS2062 (std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOSafetyDeviceForCondenserValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IS 2002(std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "IS 2062(std)", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOBlowerValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1x100% (std)", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "2x100%", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOEjectionSystemForContinuousValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Auxilary steam", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 3;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBOAutoGlandSealingSystemValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 0;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformMechanicalDBORequiredNotRequiredValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase))
            return 0;
        else if (string.Equals(str, "Not Required (std)", StringComparison.OrdinalIgnoreCase))
            return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase))
            return 2;
        else
            return isNullable ? DBNull.Value : 0;
    }

    private object TransformCurrencyValue(object value)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        var str = value.ToString()?.Trim();
        if (string.IsNullOrEmpty(str))
            return DBNull.Value;

        // Remove currency symbols and other non-numeric chars (except dot, minus and E for scientific notation)
        // Keep digits, '.', '-', 'E', 'e'
        var cleanStr = new string(str.Where(c => char.IsDigit(c) || c == '.' || c == '-' || c == 'E' || c == 'e').ToArray());

        if (decimal.TryParse(cleanStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var result))
        {
            return result;
        }

        return DBNull.Value;
    }

    #region ElectricalDBO
    private object TransformElectricalInstrumentationDBOMakeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TDPS", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "WEG", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "BHEL", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "LEROYSOMER", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "CUMMINS", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "CGL", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "Marelli Motori", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "As per Vendor List", StringComparison.OrdinalIgnoreCase)) return 8;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOStandardValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IEC 60034", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IS 4722", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "NEMA MG-1", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "API 546", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOVoltageValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "415 V", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "3300 V", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "6600 V", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "11000 V", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOEnclosureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP 23", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP 44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP 54", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP 55", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBODesignTempValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "40", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "45", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "50", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBORatedPfValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.80", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.85", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.90", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTempRiseValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Class F", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Class B", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOInsulationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Class F with VPI", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Class H", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTerminalBoxValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Fitted with CTs & N-Phase Link", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Standard", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOCertValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Safe Area", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Hazardous Area - Zone 2, Gas Gr. IIA/IIB, T3", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Hazardous Area - Zone 1, Gas Gr. IIA/IIB, T3", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBONeutralCtValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Internally Formed (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Externally Formed", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOPhaseCtValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Inside Alternator Body (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Inside Phase Terminal Box", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOOverloadValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "10% for 1 hour every 12 hours (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not applicable (for Base load)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBONoiseValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "85 dB(A) at 1 mtr (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 1;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOSlipRingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Brushless (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Static", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOPMGValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "PMG (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Aux. Winding / Arep", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTestsValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Routine test as per Std. (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Type Test as per Std", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOCoolingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IC 01 - Open Air (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IC 8A 1W7 - CACW", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IC 6A 1A1 - CACA", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOCoolerConfigValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1 x 100% (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2 x 50%", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "2 x 100%", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Cu-Ni 90/10 (Std for CACW)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Stainless Steel (Std for CACA)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOIPRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 5;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOControlModeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "PF / KVAR Control", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Voltage Control", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBusBarMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBORelayTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Numerical (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Static", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Electromagnetic", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOSyncTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Auto / Manual (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Manual Only", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOMeterAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Cl 0.5 (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Cl 0.2", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTvmTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "DLMS protocol (Std for TEDA)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Non-DLMS protocol", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOPqmValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTransformerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "NGR (Std for High/Low resistance grounding)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "NGT (Std for High resistance grounding with NGT)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOFaultRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "25 kA for 1 sec", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "25 kA for 3 sec", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "31.5 kA for 1 sec", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "31.5 kA for 3 sec", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "40 kA for 1 sec", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "40 kA for 3 sec", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "50 kA for 1 sec", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "50 kA for 3 sec", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 8;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBreakerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "VCB (Std for HT)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "ACB (Std for LT)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "SF6", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOMccConstTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Fix (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Drawout", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTcpTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "PLC Based (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Microprocessor Based", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTcpRedundancyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Dual Hot Standby (Std for Governor/Turbine Control)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Simplex (Std for protection)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "TMR", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTgpTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Local Gauge Panel (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 1;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOVmsVibrValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Contactless (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Contact", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOYessNoRequiredValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Yes", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "No", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Required (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 1;
    }

    private object TransformElectricalInstrumentationDBOTcpSpecificationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL std", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer / Consultant specification", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTcpCommunicationTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Modbus (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "RS485", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Profibus", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTvmMountingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Part MCSP panel (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Part of LASCTP panel (With separate cubical)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Separate panel", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOLtPowerCablingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL scope for TTL supplied equipments", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Purchaser scope (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBatteryCapacityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "100 AH (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 1;
        return isNullable ? DBNull.Value : 0;
    }

    //Nithya ELDBO

    private object TransformElectricalInstrumentationDBOHtPowerCablingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL scope", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Purchaser scope (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOLtPowerCableMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBusDuctValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL scope", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Purchaser scope (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOControlCableMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOHtPowerCableMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBusDuctTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "HT cables from Generator to NGR & LAPT", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "HT cables from Generator to NGR & LAPT and NGR/LAPT to Generator VCB", StringComparison.OrdinalIgnoreCase)) return 1;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBatteryTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "VRLA (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "NI-CD", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "PLANTE", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOBatteryTypeOfChargerValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1FC + 1FCBC (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2FCBC", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalInstrumentationDBOTcpSilRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "SIL-2", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "SIL-3", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMakeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TDPS", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "WEG", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "BHEL", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "LEROYSOMER", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "CUMMINS", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "CGL", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "Marelli Motori", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "As per Vendor List", StringComparison.OrdinalIgnoreCase)) return 8;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOStandardValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IEC60034(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "API546", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOVoltageAlternatorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "High Voltage Alternator", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Low Voltage Alternator", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOEnclosureValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP23(Std for LT SPD)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP23(Std for HT SPD)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP 54(Std for HT/LT CACW)", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBODesignTemperatureInDegCValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "40°C", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "45°C", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "50°C", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBORatedPfValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.8(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.9", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.85", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "0.95", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3; // Note: ID 3 is reused in JSON for Others, but 0.95 is also 3? Check JSON.
        // JSON Says: { id: 3, label: '0.95', ... }, { id: 3, label: 'Others', ... }
        // This looks like a typo in JSON, but I must follow the IDs provided if they are what the DB expects.
        // Wait, if both map to 3, how do I distinguish?
        // Assuming the DB stores the ID. I should return 3 for both.
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTemperatureRiseValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Class H(Std for SDPD", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Class B(Std for HT/LT CACW)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOInsulationClassValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Class H(Std for SDPD)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Class F(Std for HT/LT CACW)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTerminalBoxToSuitValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Cable", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Bus Duct", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAlternatorCertificationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "CE", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "GOST-R", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Not Required(TTL Standards)", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONeutralCtStarFormationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "NGR Panel(Std for HV set)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "AB Panel(Std for LV set)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Alternator Terminal Box", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Bus Duct", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOPhaseSideCtLocationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "LAPT Panel(Std for HV set)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "ACB Panel(Std for LV set)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Alternator TB", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Bus Duct", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOContinuousOverloadValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0%(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "5%", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "10%", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONoiseLevelValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "80dBA(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "85dBA", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "90dBA(Std for HV&LV alternator with CACA/CACW cooler)", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "109dBA(Std for LV SDPD alternator)", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 4;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOSlipRingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required (tender/enquiry requirements)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std for below 15MW)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Vendor standard for 15 MW & above", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOPmgValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required (tender/enquiry requirements)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std for below 15MW)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Vendor standard for 15 MW & above", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTestsValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Routine Tests (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Routine + Type tests", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCoolingMethodValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "CACA", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "SDPD", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCoolerConfigurationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Single Cooler", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Double Cooler", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCoolerTubesMocValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "SS 304", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "SS 316", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "SS 316 L", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Admiralty Brass", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Cupro Nickel", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Titanium", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 6;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCoolerMountingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Side", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Bottom", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Top", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCoolerCertificationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "PED", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Ustamping", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Not Required(TTL Standards)", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }
    private object TransformElectricalDBOAvrPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAvrTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Analog", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Digital", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Machine Mounted Without grid synchronisation", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Machine Mounted (Not applicable for grid synchronisation)", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingAvrValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOControlModeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1A + 1M", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2A + 1M", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "2A + 2M", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOPanelQtyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1No.s(std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2No.s", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOStandbyExcitationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAcbPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBusBarMaterialAcbValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Aluminum(Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAcbRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "630", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "800", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "1000", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "1250", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "1600", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "2000", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "2500", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "3200", StringComparison.OrdinalIgnoreCase)) return 8;
        else if (string.Equals(str, "4000", StringComparison.OrdinalIgnoreCase)) return 9;
        else if (string.Equals(str, "5000", StringComparison.OrdinalIgnoreCase)) return 10;
        else if (string.Equals(str, "6000", StringComparison.OrdinalIgnoreCase)) return 11;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 12;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONumberOfBreakersValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1No.(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2No.s", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONumberOfPolesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "3Poles(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "4Poles", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingAcbValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBORedundantCtPtProtectionAcbValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required(Std for redundant relays", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required(Std for non-redundant relays)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOGeneratorRelayPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBORedundantCtPtProtectionGrpValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required(Std for redundant relays", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required(Std for non-redundant relays)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBORelayTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Digital(Std for redundant relay)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Digital(Std for non-redundant relay)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingGrpValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAdditionalRelayValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Rotor earth fault", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Vector surge relay(dv/dt)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Stand by earth fault", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Overall differential protection", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Generator transformer protection", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "As per purchase order", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOSoftwareGrpValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMeteringSyncPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOSynchronizationGridValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "DG set (Momentary)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "DG set (Continuous)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Other TG sets", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTypeOfSynchronizerValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Auto", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Manual (Std for momentary synchronization)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONumberOfBreakerForSynchValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "3", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "4", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "5", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMeteringAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.5 class (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.2 class", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.2s class", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTvmTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Non-Sealed type (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Sealed type", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Sealed type with ABT feature", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTvmMountingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Part MCSP panel (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Part of LASCTP panel (With separate cubical)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Separate panel", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTvmAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.5 class (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.2 class", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.2s class", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOPqmValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTransducerQuantityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1 (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTransducerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Discrete", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Multi Function (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOLoadSharingModulesScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL scope", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Purchaser scope", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONumberOfMasterModulesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;

        // This mapping has values 0 to 8, but they correspond to '0', '1', '2'... '7', 'Others'
        if (string.Equals(str, "0", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "1", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "2", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "3", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "4", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "5", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "6", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "7", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 8;
        else return isNullable ? DBNull.Value : 0;
    }
    private object TransformElectricalDBONumberOfSlaveModulesValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "1", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "2", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "3", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "4", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "5", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "6", StringComparison.OrdinalIgnoreCase)) return 6;
        else if (string.Equals(str, "7", StringComparison.OrdinalIgnoreCase)) return 7;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 8;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMcsPanelPartOfGRPValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Yes", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "No (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOHmiSoftwareLoadSharingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    // NGR/NGT Panel Options
    private object TransformElectricalDBONgrNgtPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTypeOfPanelValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "NGR(Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "NGT", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBODutyRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "10 secs(Std)", StringComparison.OrdinalIgnoreCase)) return 0; // Note: Label in JSON says "10 secs(Std)", Value: "100%" - Wait, check JSON carefully. JSON: { id: 0, label: '10 secs(Std)', value: '100%' }? Typo in JSON value? Assuming mapping based on string equality to Label or Value. Usually we check against the string in Excel. Assuming Excel has "10 secs(Std)" or similar.
        // Actually, looking at JSON: { id: 0, label: '10 secs(Std)', value: '100%' } -> This looks like a mistake in the JSON structure provided by user. "100%" seems wrong for Duty Rating.
        // However, if the Excel contains "10 secs(Std)", it maps to 0.
        // If Excel contains "100%", it might also map to 0?
        // I will match against "10 secs(Std)" and "100%" just in case.
        if (string.Equals(str, "10 secs(Std)", StringComparison.OrdinalIgnoreCase) || string.Equals(str, "100%", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "30 secs", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTemperatureRaiseValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "375 deg.C (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "325 deg.C", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOResistorCurrentLimitingCapacityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "100 A (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBONeutralIsolatorValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Manual (Std if without grid synchronisation)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Motorised (Std if with grid synchronisation)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCtValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Non Redundant (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Redundant", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingNGRTValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMeteringCtAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.5 class (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.2 class", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.2S class", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOFaultRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "26 KA (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "40 KA", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBusBarMaterialValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    // LASC/PT Panel Options
    private object TransformElectricalDBOLascptPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCtPtValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Yes", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "No (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCtPtAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.5 class (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.2 class", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.2S class", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOFaultRatingLascptValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "26 KA (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "40 KA", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingLascptValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBusBarMaterialLascptValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOLighteningArrestorSurgeCapacitorPtPanelPartOfBreakerPanelValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Yes", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "No (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }
    // Switch Gear Panel Options
    private object TransformElectricalDBOSwitchGearPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOQuantityAndRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "As per enclosed SLD", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Standard", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBreakerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "ACB (Std for LT set)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "VCB (Std) for HT set", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "SF6", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOMeteringCtPtAccuracyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "0.5 class (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "0.2 class", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "0.2s class", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOFaultRatingSgValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "26 KA (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "40 KA", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingSgValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBusBarMaterialSgValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    // MCC Panel Options
    private object TransformElectricalDBOMccScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOStandByExcitationTfValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIncomerQtyValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1 (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "2 + Bus Coupler", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOTypeOfConstructionValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Draw-out type", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Non draw-out type (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBORedundantControlTfValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Required", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Not Required (Std)", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOSpecificationValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer Specification", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIncomerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "MCCB", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "MPCB", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "ACB", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "SFU (Std)", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBusBarMaterialMccValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Aluminium (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Copper", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingMccValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOAcdbValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Part of MCC panel (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Separate panel", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    // Battery Charger Panel Options
    private object TransformElectricalDBOBatteryChargerPanelScopeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "TTL", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Customer", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Existing", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Not Applicable", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBODcdbValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Part of B&BC Panel (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "Separate Panel", StringComparison.OrdinalIgnoreCase)) return 1;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOVoltageRatingValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "24 V (Without EOP)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "110 V (Std with EOP)", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "220 V", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOCapacityValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "100 AH", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "200 AH", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "400 AH", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "600 AH", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "800 AH", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOBatteryTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "VRLA (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "NI-CD", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "PLANTE", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 3;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOChargerTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "1FC + 1FCBC (Std)", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "2FCBC", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "Others", StringComparison.OrdinalIgnoreCase)) return 2;
        else return isNullable ? DBNull.Value : 0;
    }

    private object TransformElectricalDBOIpRatingBatteryValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "IP-42", StringComparison.OrdinalIgnoreCase)) return 0;
        else if (string.Equals(str, "IP-44", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "IP-52", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "IP-54", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "IP-55", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Std", StringComparison.OrdinalIgnoreCase)) return 5;
        else return isNullable ? DBNull.Value : 0;
    }

    #endregion ElectricalDBO
    private object TransformMOMIsPresentValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
        {
            return isNullable ? DBNull.Value : 0;
        }

        var stringValue = value.ToString()?.Trim().ToLowerInvariant();

        switch (stringValue)
        {
            case "yes":
            case "y":
            case "present":
            case "1":
                return 1;
            default:
                return 0;
        }
    }

    private object TransformMOMMeetingTypeValue(object value, bool isNullable)
    {
        if (value == null || value == DBNull.Value)
        {
            return isNullable ? DBNull.Value : 0;
        }

        var str = value.ToString()?.Trim() ?? string.Empty;
        if (string.Equals(str, "Select", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(str)) return 0;
        else if (string.Equals(str, "Internal", StringComparison.OrdinalIgnoreCase)) return 1;
        else if (string.Equals(str, "With Customer", StringComparison.OrdinalIgnoreCase) || string.Equals(str, "WithCustomer", StringComparison.OrdinalIgnoreCase)) return 2;
        else if (string.Equals(str, "Safety", StringComparison.OrdinalIgnoreCase)) return 3;
        else if (string.Equals(str, "Weekly Status", StringComparison.OrdinalIgnoreCase) || string.Equals(str, "WeeklyStatus", StringComparison.OrdinalIgnoreCase)) return 4;
        else if (string.Equals(str, "Monthly Status", StringComparison.OrdinalIgnoreCase) || string.Equals(str, "MonthlyStatus", StringComparison.OrdinalIgnoreCase)) return 5;
        else if (string.Equals(str, "Other", StringComparison.OrdinalIgnoreCase)) return 6;

        return isNullable ? DBNull.Value : 0;
    }

    private object ConvertValue(object value, Type targetType)
    {
        if (value == null || value == DBNull.Value)
            return DBNull.Value;

        // If types match, return as-is
        if (value.GetType() == targetType || targetType.IsAssignableFrom(value.GetType()))
            return value;

        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

        // Check if value is a string that might contain scientific notation
        var valueString = value as string ?? value.ToString() ?? string.Empty;
        var isScientificNotation = !string.IsNullOrWhiteSpace(valueString) &&
            (valueString.Contains('E', StringComparison.OrdinalIgnoreCase) ||
             valueString.Contains('e', StringComparison.OrdinalIgnoreCase));

        // Handle numeric conversions with scientific notation support
        if (underlyingType == typeof(int))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return (int)double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToInt32(value);
        }
        if (underlyingType == typeof(long))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return (long)double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToInt64(value);
        }
        if (underlyingType == typeof(short))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return (short)double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToInt16(value);
        }
        if (underlyingType == typeof(byte))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return (byte)double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToByte(value);
        }
        if (underlyingType == typeof(decimal))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                // Parse as double first, then convert to decimal to handle scientific notation
                var doubleValue = double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                return (decimal)doubleValue;
            }
            // Try parsing as decimal, but if it fails and contains 'E', try scientific notation
            try
            {
                return Convert.ToDecimal(value);
            }
            catch
            {
                // If standard conversion fails, try parsing with scientific notation support
                if (!string.IsNullOrWhiteSpace(valueString))
                {
                    var doubleValue = double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                    return (decimal)doubleValue;
                }
                return Convert.ToDecimal(value);
            }
        }
        if (underlyingType == typeof(double))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToDouble(value);
        }
        if (underlyingType == typeof(float))
        {
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString))
            {
                return (float)double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
            return Convert.ToSingle(value);
        }
        if (underlyingType == typeof(bool))
        {
            // Handle text-based boolean values (Yes/No, Required/Not Required, etc.)
            if (value is string stringValue && !string.IsNullOrWhiteSpace(stringValue))
            {
                var trimmedValue = stringValue.Trim();

                // Check for "true" values (should convert to 1/true)
                if (string.Equals(trimmedValue, "Yes", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "Y", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "True", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "1", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "Required", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                // Check for "false" values (should convert to 0/false)
                if (string.Equals(trimmedValue, "No", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "N", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "False", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "0", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "Not Required", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmedValue, "NotRequired", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            // Try standard conversion for other cases
            return Convert.ToBoolean(value);
        }
        if (underlyingType == typeof(DateTime))
        {
            return Convert.ToDateTime(value);
        }
        if (underlyingType == typeof(string))
        {
            return value.ToString() ?? string.Empty;
        }

        // Try direct conversion
        try
        {
            return Convert.ChangeType(value, underlyingType);
        }
        catch
        {
            // If direct conversion fails and value is a string with scientific notation, try parsing as double first
            if (isScientificNotation && !string.IsNullOrWhiteSpace(valueString) && (underlyingType == typeof(decimal) || underlyingType == typeof(double) || underlyingType == typeof(float)))
            {
                var doubleValue = double.Parse(valueString, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                return Convert.ChangeType(doubleValue, underlyingType);
            }
            throw;
        }
    }

    private async Task<int> BulkCopyToTempTableAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string tempTableName,
        DataTable dataTable,
        List<ColumnMapping> mappings,
        bool hasIdentityInExcel,
        CancellationToken cancellationToken)
    {
        if (dataTable.Rows.Count == 0)
            return 0;

        var options = SqlBulkCopyOptions.Default;
        if (hasIdentityInExcel)
        {
            options |= SqlBulkCopyOptions.KeepIdentity;
        }

        using var bulkCopy = new SqlBulkCopy(connection, options, transaction);
        bulkCopy.DestinationTableName = tempTableName;
        bulkCopy.BulkCopyTimeout = SqlCommandTimeout; // 10 minutes for large datasets

        // Map columns
        foreach (var mapping in mappings)
        {
            bulkCopy.ColumnMappings.Add(mapping.SqlColumnName, mapping.SqlColumnName);
        }

        await bulkCopy.WriteToServerAsync(dataTable, cancellationToken);

        return dataTable.Rows.Count;
    }

    private async Task<int> InsertFromTempToTargetAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string tempTableName,
        List<ColumnMapping> mappings,
        ColumnMetadata? identityColumn,
        bool hasIdentityInExcel,
        CancellationToken cancellationToken)
    {
        var hasIsDeletedColumn = await CheckColumnExistsAsync(connection, transaction, schemaName, tableName, "IsDeleted", cancellationToken);
        var enableIdentityInsert = hasIdentityInExcel && identityColumn != null;

        var insertMappings = mappings
            .Where(m => !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var columnList = string.Join(", ", insertMappings.Select(m => $"[{m.SqlColumnName}]"));
        var selectList = string.Join(", ", insertMappings.Select(m => $"source.[{m.SqlColumnName}]"));

        if (hasIsDeletedColumn)
        {
            columnList = columnList + ", [IsDeleted]";
            selectList = selectList + ", 0";
        }

        var insertQuery = $"INSERT INTO [{schemaName}].[{tableName}] ({columnList}) SELECT {selectList} FROM {tempTableName} AS source";

        try
        {
            if (enableIdentityInsert)
            {
                var enableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] ON";
                await using var cmd1 = new SqlCommand(enableIdentityCmd, connection, transaction);
                cmd1.CommandTimeout = SqlCommandTimeout;
                await cmd1.ExecuteNonQueryAsync(cancellationToken);
            }

            await using var command = new SqlCommand(insertQuery, connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            return await command.ExecuteNonQueryAsync(cancellationToken);
        }
        finally
        {
            if (enableIdentityInsert)
            {
                var disableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] OFF";
                await using var cmd2 = new SqlCommand(disableIdentityCmd, connection, transaction);
                cmd2.CommandTimeout = SqlCommandTimeout;
                await cmd2.ExecuteNonQueryAsync(cancellationToken);
            }
        }
    }

    private async Task<(int rowsInserted, int rowsUpdated)> MergeFromTempToTargetAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string tempTableName,
        List<ColumnMapping> mappings,
        List<string> primaryKeyColumns,
        ColumnMetadata? identityColumn,
        bool hasIdentityInExcel,
        CancellationToken cancellationToken)
    {
        // Check if IsDeleted column exists in target table
        var hasIsDeletedColumn = await CheckColumnExistsAsync(connection, transaction, schemaName, tableName, "IsDeleted", cancellationToken);
        return await MergeFromTempToTargetAsyncInternal(connection, transaction, schemaName, tableName, tempTableName, mappings, primaryKeyColumns, identityColumn, hasIdentityInExcel, hasIsDeletedColumn, cancellationToken);
    }

    private async Task<bool> CheckColumnExistsAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string columnName,
        CancellationToken cancellationToken)
    {
        var query = @"
            SELECT COUNT(*)
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = @SchemaName
                AND TABLE_NAME = @TableName
                AND COLUMN_NAME = @ColumnName";

        await using var command = new SqlCommand(query, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        command.Parameters.AddWithValue("@SchemaName", schemaName);
        command.Parameters.AddWithValue("@TableName", tableName);
        command.Parameters.AddWithValue("@ColumnName", columnName);

        var result = await command.ExecuteScalarAsync(cancellationToken);
        return result != null && Convert.ToInt32(result) > 0;
    }

    private async Task<(int rowsInserted, int rowsUpdated)> MergeFromTempToTargetAsyncInternal(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string tempTableName,
        List<ColumnMapping> mappings,
        List<string> primaryKeyColumns,
        ColumnMetadata? identityColumn,
        bool hasIdentityInExcel,
        bool hasIsDeletedColumn,
        CancellationToken cancellationToken)
    {
        var identityColumnName = identityColumn?.ColumnName;
        var enableIdentityInsert = hasIdentityInExcel && identityColumn != null;

        // If no primary key exists, fall back to INSERT (with potential duplicates)
        if (primaryKeyColumns.Count == 0)
        {
            // Fallback to simple INSERT if no primary key
            // Exclude IsDeleted from regular mappings and handle it separately
            var insertMappingsNoPk = mappings
                .Where(m => !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var columnList = string.Join(", ", insertMappingsNoPk.Select(m => $"[{m.SqlColumnName}]"));
            var selectList = string.Join(", ", insertMappingsNoPk.Select(m => $"source.[{m.SqlColumnName}]"));

            // Always set IsDeleted to 0/false if the column exists
            if (hasIsDeletedColumn)
            {
                columnList = columnList + ", [IsDeleted]";
                selectList = selectList + ", 0";
            }

            var insertQuery = $"INSERT INTO [{schemaName}].[{tableName}] ({columnList}) SELECT {selectList} FROM {tempTableName} AS source";

            try
            {
                if (enableIdentityInsert)
                {
                    var enableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] ON";
                    await using var cmd1 = new SqlCommand(enableIdentityCmd, connection, transaction);
                    cmd1.CommandTimeout = SqlCommandTimeout;
                    await cmd1.ExecuteNonQueryAsync(cancellationToken);
                }

                await using var command = new SqlCommand(insertQuery, connection, transaction);
                command.CommandTimeout = SqlCommandTimeout;
                var rowsAffected = await command.ExecuteNonQueryAsync(cancellationToken);
                return (rowsAffected, 0);
            }
            finally
            {
                if (enableIdentityInsert)
                {
                    var disableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] OFF";
                    await using var cmd2 = new SqlCommand(disableIdentityCmd, connection, transaction);
                    cmd2.CommandTimeout = SqlCommandTimeout;
                    await cmd2.ExecuteNonQueryAsync(cancellationToken);
                }
            }
        }

        // Build MERGE statement for UPSERT
        var sb = new StringBuilder();

        // Create table variable to capture merge results
        sb.AppendLine("DECLARE @MergeResults TABLE (Action NVARCHAR(10));");
        sb.AppendLine();

        sb.AppendLine($";MERGE [{schemaName}].[{tableName}] AS target");
        sb.AppendLine($"USING {tempTableName} AS source");

        // Build ON clause for primary key matching
        // Handle NULL values properly: use ISNULL to convert NULL to a sentinel value for comparison
        // This ensures that NULL = NULL evaluates to TRUE in the MERGE ON clause
        var matchConditions = primaryKeyColumns
            .Where(pk => mappings.Any(m => m.SqlColumnName.Equals(pk, StringComparison.OrdinalIgnoreCase)))
            .Select(pk =>
            {
                // For proper NULL handling, use a pattern that works for all data types
                // Use COALESCE with a type-appropriate sentinel value, or use ISNULL
                // For numeric types, use -1 or 0; for strings, use empty string; for dates, use a far future date
                // However, since PKs typically can't be NULL, we'll use a simpler approach:
                // Use ISNULL to handle potential NULLs, but also ensure exact matching
                return $"(target.[{pk}] = source.[{pk}] OR (target.[{pk}] IS NULL AND source.[{pk}] IS NULL))";
            });

        if (!matchConditions.Any())
        {
            // If primary key columns are not in mappings, fall back to INSERT
            // Exclude IsDeleted from regular mappings and handle it separately
            var insertMappingsFallback = mappings
                .Where(m => !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var columnList = string.Join(", ", insertMappingsFallback.Select(m => $"[{m.SqlColumnName}]"));
            var selectList = string.Join(", ", insertMappingsFallback.Select(m => $"source.[{m.SqlColumnName}]"));

            // Always set IsDeleted to 0/false if the column exists
            if (hasIsDeletedColumn)
            {
                columnList = columnList + ", [IsDeleted]";
                selectList = selectList + ", 0";
            }

            var insertQuery = $"INSERT INTO [{schemaName}].[{tableName}] ({columnList}) SELECT {selectList} FROM {tempTableName} AS source";

            try
            {
                if (enableIdentityInsert)
                {
                    var enableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] ON";
                    await using var cmd1 = new SqlCommand(enableIdentityCmd, connection, transaction);
                    cmd1.CommandTimeout = SqlCommandTimeout;
                    await cmd1.ExecuteNonQueryAsync(cancellationToken);
                }

                await using var command = new SqlCommand(insertQuery, connection, transaction);
                command.CommandTimeout = SqlCommandTimeout;
                var rowsAffected = await command.ExecuteNonQueryAsync(cancellationToken);
                return (rowsAffected, 0);
            }
            finally
            {
                if (enableIdentityInsert)
                {
                    var disableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] OFF";
                    await using var cmd2 = new SqlCommand(disableIdentityCmd, connection, transaction);
                    cmd2.CommandTimeout = SqlCommandTimeout;
                    await cmd2.ExecuteNonQueryAsync(cancellationToken);
                }
            }
        }

        sb.AppendLine($"ON {string.Join(" AND ", matchConditions)}");

        // Build WHEN MATCHED clause - UPDATE all non-PK columns
        // Get all columns that are in mappings but NOT in primary key and NOT IsDeleted
        var nonPkColumns = mappings
            .Where(m => !primaryKeyColumns.Any(pk => pk.Equals(m.SqlColumnName, StringComparison.OrdinalIgnoreCase)) &&
                       !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Ensure we have columns to update
        if (nonPkColumns.Count == 0)
        {
            // If all columns are primary keys, we can't update anything
            // This shouldn't happen in practice, but handle it gracefully
            nonPkColumns = mappings
                .Where(m => !m.IsIdentity && !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        // Always include WHEN MATCHED clause to update all non-PK columns
        // Ensure we update even if values are the same (MERGE will still count as UPDATE)
        var updateClausesList = new List<string>();

        if (nonPkColumns.Any())
        {
            // Update ALL non-PK columns from source to target (excluding IsDeleted)
            updateClausesList.AddRange(nonPkColumns.Select(m => $"[{m.SqlColumnName}] = source.[{m.SqlColumnName}]"));
        }
        else
        {
            // If somehow no non-PK columns, update all columns except PK and IsDeleted (fallback)
            var allUpdateColumns = mappings
                .Where(m => !primaryKeyColumns.Contains(m.SqlColumnName, StringComparer.OrdinalIgnoreCase) &&
                           !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
                .Select(m => $"[{m.SqlColumnName}] = source.[{m.SqlColumnName}]");

            if (allUpdateColumns.Any())
            {
                // Fallback to all columns except PK
                updateClausesList.AddRange(allUpdateColumns);
            }
        }

        // Always set IsDeleted to 0/false if the column exists
        if (hasIsDeletedColumn)
        {
            // Check if IsDeleted is already in the list
            if (!updateClausesList.Any(c => c.Contains("[IsDeleted]", StringComparison.OrdinalIgnoreCase)))
            {
                updateClausesList.Add("[IsDeleted] = 0");
            }
        }

        // If no columns to update, try to find a timestamp column as last resort
        if (!updateClausesList.Any())
        {
            var timestampColumn = mappings.FirstOrDefault(m =>
                m.SqlColumnName.EndsWith("UpdatedAt", StringComparison.OrdinalIgnoreCase) ||
                m.SqlColumnName.EndsWith("LastUpdated", StringComparison.OrdinalIgnoreCase) ||
                m.SqlColumnName.EndsWith("ModifiedDate", StringComparison.OrdinalIgnoreCase));

            if (timestampColumn != null)
            {
                updateClausesList.Add($"[{timestampColumn.SqlColumnName}] = ISNULL(source.[{timestampColumn.SqlColumnName}], GETDATE())");
            }
        }

        // Only add WHEN MATCHED clause if we have something to update
        if (updateClausesList.Any())
        {
            sb.AppendLine("WHEN MATCHED THEN");
            sb.AppendLine("UPDATE SET");
            sb.AppendLine(string.Join(",\n", updateClausesList));
        }

        // Build WHEN NOT MATCHED clause - INSERT new rows
        sb.AppendLine("WHEN NOT MATCHED BY TARGET THEN");

        // Exclude IsDeleted from regular mappings and handle it separately
        var insertMappings = mappings
            .Where(m => !string.Equals(m.SqlColumnName, "IsDeleted", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var insertColumnsList = insertMappings.Select(m => $"[{m.SqlColumnName}]").ToList();
        var insertValuesList = insertMappings.Select(m => $"source.[{m.SqlColumnName}]").ToList();

        // Always set IsDeleted to 0/false if the column exists
        if (hasIsDeletedColumn)
        {
            insertColumnsList.Add("[IsDeleted]");
            insertValuesList.Add("0");
        }

        var insertColumns = string.Join(", ", insertColumnsList);
        var insertValues = string.Join(", ", insertValuesList);
        sb.AppendLine($"INSERT ({insertColumns}) VALUES ({insertValues})");

        // Output clause to track inserted/updated rows
        sb.AppendLine("OUTPUT $action INTO @MergeResults;");
        sb.AppendLine();
        sb.AppendLine("SELECT ");
        sb.AppendLine("    SUM(CASE WHEN Action = 'INSERT' THEN 1 ELSE 0 END) AS InsertedCount,");
        sb.AppendLine("    SUM(CASE WHEN Action = 'UPDATE' THEN 1 ELSE 0 END) AS UpdatedCount");
        sb.AppendLine("FROM @MergeResults;");

        try
        {
            if (enableIdentityInsert)
            {
                var enableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] ON";
                await using var cmd1 = new SqlCommand(enableIdentityCmd, connection, transaction);
                cmd1.CommandTimeout = SqlCommandTimeout;
                await cmd1.ExecuteNonQueryAsync(cancellationToken);
            }

            await using var command = new SqlCommand(sb.ToString(), connection, transaction);
            command.CommandTimeout = SqlCommandTimeout;
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);

            int insertedCount = 0;
            int updatedCount = 0;

            // The MERGE statement produces one result set (the OUTPUT into table variable)
            // Then the SELECT statement produces the second result set with counts
            // We need to skip the first result set (MERGE OUTPUT) and read the second (SELECT counts)
            // Note: If no rows were processed, @MergeResults will be empty and SUM will return NULL
            if (await reader.NextResultAsync(cancellationToken))
            {
                if (await reader.ReadAsync(cancellationToken))
                {
                    // Handle NULL values from SUM (when no rows processed)
                    insertedCount = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                    updatedCount = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                }
            }
            else
            {
                // If NextResultAsync returns false, there's no second result set
                // This shouldn't happen, but handle it gracefully
                // The MERGE should have executed, so we'll return 0,0
                // In practice, this might indicate an error, but we'll let the transaction handle it
            }

            return (insertedCount, updatedCount);
        }
        finally
        {
            if (enableIdentityInsert)
            {
                var disableIdentityCmd = $"SET IDENTITY_INSERT [{schemaName}].[{tableName}] OFF";
                await using var cmd2 = new SqlCommand(disableIdentityCmd, connection, transaction);
                cmd2.CommandTimeout = SqlCommandTimeout;
                await cmd2.ExecuteNonQueryAsync(cancellationToken);
            }
        }
    }

    private async Task DropTempTableAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string tempTableName,
        CancellationToken cancellationToken)
    {
        var dropQuery = $"IF OBJECT_ID('tempdb..{tempTableName}') IS NOT NULL DROP TABLE {tempTableName}";
        await using var command = new SqlCommand(dropQuery, connection, transaction);
        command.CommandTimeout = SqlCommandTimeout;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private class ColumnMetadata
    {
        public string ColumnName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public bool IsNullable { get; set; }
        public int? MaxLength { get; set; }
        public int? NumericPrecision { get; set; }
        public int? NumericScale { get; set; }
        public bool IsIdentity { get; set; }
        public bool IsPrimaryKey { get; set; }
        public string? ForeignKeyTableSchema { get; set; }
        public string? ForeignKeyTableName { get; set; }
        public string? ForeignKeyColumnName { get; set; }
        public string? ForeignKeyLookupColumnName { get; set; } // Column in parent table to search by (usually a name/description column)
    }

    private class ColumnMapping
    {
        public string ExcelColumnName { get; set; } = string.Empty;
        public string SqlColumnName { get; set; } = string.Empty;
        public string SqlDataType { get; set; } = string.Empty;
        public bool IsIdentity { get; set; }
        public bool IsNullable { get; set; }
        public string? ForeignKeyTableSchema { get; set; }
        public string? ForeignKeyTableName { get; set; }
        public string? ForeignKeyColumnName { get; set; }
        public string? ForeignKeyLookupColumnName { get; set; }
    }



    private async Task<bool> RecordExistsAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        string schemaName,
        string tableName,
        string idColumn,
        string idValue,
        CancellationToken cancellationToken)
    {
        // Try to parse the ID as a long
        if (!long.TryParse(idValue, out long parsedId))
        {
            return false; // Invalid ID format (non-numeric), so it doesn't exist
        }

        var query = $"SELECT COUNT(1) FROM [{schemaName}].[{tableName}] WHERE [{idColumn}] = @Id";
        await using var command = new SqlCommand(query, connection, transaction);
        command.Parameters.AddWithValue("@Id", parsedId);
        command.CommandTimeout = SqlCommandTimeout;

        var count = (int?)(await command.ExecuteScalarAsync(cancellationToken)) ?? 0;
        return count > 0;
    }

    private async Task<UploadResponse> MigrateToRCCALineItemsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        // Specific filtering logic: uuu_tab_id = 0 for RCCA_StandardLI
        DataTable? filteredData = null;
        if (excelData.Columns.Contains("uuu_tab_id"))
        {
            string targetTabId = "0";
            if (string.Equals(tableName, "RCCA_SelectTeamMembersLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "1";
            }

            var rows = excelData.AsEnumerable()
                .Where(row => row.Field<object>("uuu_tab_id")?.ToString()?.Trim() == targetTabId);

            if (rows.Any())
            {
                filteredData = rows.CopyToDataTable();
            }
        }
        else
        {
            filteredData = excelData;
        }

        if (filteredData == null || filteredData.Rows.Count == 0)
        {
            response.Success = true;
            response.Message = $"No matching rows found with uuu_tab_id = 0 for table {tableName}. No data migrated.";
            return response;
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        return await MigrateToSingleTableAsync(
            connection,
            schemaName,
            tableName,
            filteredData,
            null,
            cancellationToken);
    }

    private async Task<UploadResponse> MigrateToMonthlyProgressReportLineItemsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        // Specific filtering logic: uuu_tab_id based on table name
        DataTable? filteredData = null;
        string targetTabId = "0";
        if (excelData.Columns.Contains("uuu_tab_id"))
        {
            if (tableName.EndsWith("ScopeOfSupplyLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "1";
            }
            else if (tableName.EndsWith("ProcurementProgressofBoughtOutsItemLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "2";
            }
            else if (tableName.EndsWith("TurbineManufacturingProgressLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "3";
            }
            else if (tableName.EndsWith("InspectionDispatchPlanLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "4";
            }
            else if (tableName.EndsWith("CashInFlowPlanLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "5";
            }
            else if (tableName.EndsWith("LookAheadTaskforNext30DaysLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "6";
            }
            else if (tableName.EndsWith("EngineeringProgressLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "7";
            }
            else if (tableName.EndsWith("InputsRequiredFromCustomerLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "8";
            }
            else if (tableName.EndsWith("DBOSummarizeSheetLI", StringComparison.OrdinalIgnoreCase))
            {
                targetTabId = "10";
            }

            var rows = excelData.AsEnumerable()
                .Where(row => row.Field<object>("uuu_tab_id")?.ToString()?.Trim() == targetTabId);

            if (rows.Any())
            {
                filteredData = rows.CopyToDataTable();
            }
        }
        else
        {
            // If column is missing, we don't skip – let it proceed but it might fail later or migrate nothing
            filteredData = excelData;
        }

        if (filteredData == null || filteredData.Rows.Count == 0)
        {
            response.Success = true;
            response.Message = $"No matching rows found with uuu_tab_id = {targetTabId} for table {tableName}. No data migrated.";
            return response;
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        return await MigrateToSingleTableAsync(
            connection,
            schemaName,
            tableName,
            filteredData,
            null,
            cancellationToken);
    }

    private bool IsValueZero(object? value)
    {
        if (value == null || value == DBNull.Value) return false;
        var str = value.ToString()?.Trim();
        return str == "0" || str == "0.0" || str == "0.00";
    }

    private async Task<UploadResponse> MigrateToOrderTransmittalNotesAsync(
        string connectionString,
        string schemaName,
        string tableName,
        DataTable excelData,
        CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        // Specific filtering logic: uuu_tab_id == 0
        DataTable? filteredData = null;
        string targetTabId = "0";
        if (excelData.Columns.Contains("uuu_tab_id"))
        {
            var rows = excelData.AsEnumerable()
                .Where(row => row.Field<object>("uuu_tab_id")?.ToString()?.Trim() == targetTabId);

            if (rows.Any())
            {
                filteredData = rows.CopyToDataTable();
            }
        }
        else
        {
            filteredData = excelData;
        }

        if (filteredData == null || filteredData.Rows.Count == 0)
        {
            response.Success = true;
            response.Message = $"No matching rows found with uuu_tab_id = {targetTabId} for table {tableName}. No data migrated.";
            return response;
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        return await MigrateToSingleTableAsync(
            connection,
            schemaName,
            tableName,
            filteredData,
            null,
            cancellationToken);
    }


}

