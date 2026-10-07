"""SQL for the three live-database pulls used by the Accenture Closed-with-Insuff report.

Queries go through pymysql's %-style parameter binding, so any literal percent
sign in the SQL text is written as %%. `{in_list}` is replaced with a run of
%s placeholders (one per value) before execution.
"""

DAILY_QUERY = """
SELECT ecc.case_id,ecc.case_check_id,emc.company_name,process_name,office_name location,
case_ars_no,CONCAT(first_name,' ',Middle_name,' ',last_name) AS Candidate_name,
received_date case_received_date,ecm.created_date AS case_created_date,
last_insuff_date,last_insuff_fulfill_date,ecff.case_flex_field1,ecff.case_flex_field2,
ecff.case_flex_field3,ecff.case_flex_field4,ecff.case_flex_field5,ecff.case_flex_field6,
ecff.case_flex_field7,ecff.case_flex_field8,ecff.case_flex_field9,ecff.case_flex_field10,
ecff.case_flex_field11,ecff.case_flex_field12,ecff.case_flex_field13,ecff.case_flex_field14,
ecff.case_flex_field15,ecff.case_flex_field16,ecff.case_flex_field17,ecff.case_flex_field18,
ecff.case_flex_field19,ecff.case_flex_field20,
case_expected_closure_date,
(CASE case_status
WHEN '1' THEN 'New (Incomplete)'
WHEN '2' THEN 'On Hold'
WHEN '3' THEN 'Insufficient'
WHEN '4' THEN 'Work in Progress'
WHEN '5' THEN 'Pending for report'
WHEN '6' THEN 'Closed by Client'
WHEN '7' THEN 'Completed'
WHEN '8' THEN 'Closed by Authbridge'
WHEN '9' THEN 'Closed - Case Insufficient' END) AS 'Case_status',
check_name,
check_created_on,
go_ahead_date,
check_expected_closure,
Insuff_date,
ecc.insuff_remarks,
Insuff_fulfill_date,
checkpoint_live.fn_check_status(check_status) AS 'check_status',
check_severity,check_closure_date,
CONCAT(eud.user_first_name,' ',eud.User_last_name) Created_by,eud1.user_id,
CONCAT(eud1.user_first_name,' ',eud1.User_last_name) Verifier,
emc1.company_name AS emp_check_verification_source,
CONCAT(emei.institute_name,' - ',emei1.institute_name) AS edu_check_verification_source,
(CASE VER_OVERSEAS
	WHEN '0' THEN 'No'
	WHEN '1' THEN 'Yes'
END) AS Ver_Overseas,
checkpoint_live.fn_ver_summary(ecc.VER_SUMMARY) AS 'Ver_Summary',
(CASE VER_TYPE
	WHEN '1' THEN 'Verbal'
	WHEN '2' THEN 'Written'
	WHEN '3' THEN 'VRWP'
END) AS Ver_Type,
checkpoint_live.fn_ver_procedure(ecc.VERIFICATION_PROCEDURE) as 'VERIFICATION_PROCEDURE',
checkpoint_live.fn_vs_name(ecc.case_check_id) as 'VS',
ecc.closure_comments as 'Closure Comments',
CONCAT(eud2.user_first_name,' ',eud2.User_last_name) Check_completed_by,
eud2.user_employee_id


FROM ec_case_checks ecc
LEFT JOIN ec_case_master ecm ON ecc.case_id=ecm.case_id
LEFT JOIN ec_master_company emc ON ecm.client_id=emc.company_id
LEFT JOIN ec_user_details eud ON ecc.documented_by=eud.user_id
LEFT JOIN ec_user_details eud1 ON ecc.check_verifier=eud1.user_id
LEFT JOIN ec_user_details eud2 ON ecc.CHECK_COMPLETED_BY_ID=eud2.user_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_educational_institute emei1 ON emei.university_id=emei1.institute_id
LEFT JOIN ec_case_fields ecff ON ecm.case_id=ecff.case_id
WHERE check_status IN (8,10,11,12) AND case_ars_no LIKE '%%-%%'
AND ecc.check_closure_date>= %s AND ecc.check_closure_date<= %s
AND emc.company_name LIKE %s
"""

ADVANCE_QUERY = """
SELECT ecc.Case_Check_id,case_ars_no,emc.Company_name,
CONCAT(first_name,' ',IFNULL(middle_name,''),' ',IFNULL(last_name,'')) AS Candidate_name,
Process_name,  received_date AS case_received_date,
ecm.created_date AS case_created_date,
ecff.CASE_FLEX_FIELD1,ecff.CASE_FLEX_FIELD2,ecff.CASE_FLEX_FIELD3,ecff.CASE_FLEX_FIELD4,
ecff.CASE_FLEX_FIELD5,ecff.CASE_FLEX_FIELD6,ecff.CASE_FLEX_FIELD7,ecff.CASE_FLEX_FIELD8,
ecff.CASE_FLEX_FIELD9,ecff.CASE_FLEX_FIELD10,ecff.CASE_FLEX_FIELD11,ecff.CASE_FLEX_FIELD12,
ecff.CASE_FLEX_FIELD13,ecff.CASE_FLEX_FIELD14,ecff.CASE_FLEX_FIELD15,ecff.CASE_FLEX_FIELD16,
ecff.CASE_FLEX_FIELD17,ecff.CASE_FLEX_FIELD18,ecff.CASE_FLEX_FIELD19,ecff.CASE_FLEX_FIELD20,
ecff.CASE_FLEX_FIELD21,ecff.CASE_FLEX_FIELD22,ecff.CASE_FLEX_FIELD23,ecff.CASE_FLEX_FIELD24,
ecff.CASE_FLEX_FIELD25,ecff.CASE_FLEX_FIELD26,ecff.CASE_FLEX_FIELD27,ecff.CASE_FLEX_FIELD28,
ecff.CASE_FLEX_FIELD29,ecff.CASE_FLEX_FIELD30,
checkpoint_live.fn_case_status(case_status) AS 'Case_status',
checkpoint_live.fn_check_status(check_status) AS 'check_status',
ec.check_name AS 'Check_unique_name',
check_disposition_id,disposition_name,check_severity,
REPLACE(ecc.closure_comments,'rn','') AS closure_comments,ecc.check_closure_date,
ecc.insuff_remarks,
-- ecc.insuff_date,

@insuffdate:=(SELECT action_taken_on FROM ec_case_history ech WHERE ecc.case_check_id=ech.check_id
AND action_taken in (' case Status changed to : Case Insufficient',
' case Status changed to : InSufficient',
'Insuff Raised',
'Marked Insufficient',
'Marked Insufficient (Parallel Research)',
'New Status - Case Insufficient',
'New Status - InSufficient',
'Vendor request closed & Insuff raised',
'Case Insufficient',
'Check Created | Marked Insufficient',
'Check Updated | Marked Insufficient',
'New Status - Case Insuff Updated',
'New Status - New Case | Case Insufficient',
'Check Insuff raised',
'Case level Insuff raised',
'New Status - Case Insufficiency raised',
'Insufficient - Rework on Report',
'Insuff accepted' ) ORDER BY action_id LIMIT 1) AS 'First_Insuff_Date',

ecc.insuff_fulfill_date,
office_name location,
IF(ecc.family_id=4,emei.institute_name,IF(ecc.family_id=5,emc1.company_name,emcy.city_name)) AS verification_source,
ecm.case_expected_closure_date case_due_date,check_created_on,ecc.go_ahead_date,ecc.copy_of_check,

IFNULL(ec.CHECK_OPS_NAME,LEFT(REPLACE(family_name,' Family',''),3)) AS 'Check Ops Name',

ecc.reopen_date AS 'check_reopen_date',ecm.reopen_date AS 'case_reopen_date',
checkpoint_live.fn_ver_summary(ecc.VER_SUMMARY) AS Ver_Summary,
checkpoint_live.fn_ver_procedure(ecc.VERIFICATION_PROCEDURE) AS VERIFICATION_PROCEDURE,
checkpoint_live.fn_check_sub_status(ecc.sub_status) AS 'Check Sub Status',
(SELECT action_taken_on FROM ec_case_history ech
WHERE ecc.case_check_id=ech.check_id AND action_taken='Insuff Qc Status Updateas'
AND action_comments='Accepted' ORDER BY action_id DESC LIMIT 1) AS 'Inusff Accepted',

(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Severity',
(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Final_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Severity',
dqc_released_date,
(CASE TIER
WHEN 0 THEN 'Overseas'
WHEN 1 THEN 'Tier 1'
WHEN 2 THEN 'Tier 2'
WHEN 3 THEN 'Tier 3'
WHEN 4 THEN 'Tier 4' END ) AS 'Tier',

CONCAT(eud1.user_first_name,' ',eud1.user_last_name) AS 'DS Name',
QUEUE_NAME,
IF(@insuffdate is not null,if(@insuffdate<=dqc_released_date,'L1','L2'),'') AS 'Insuff Type',
IF(@insuffdate is not null,if(check_status in (0,1,2,3,4,5,6,7,13),'WIP','Non-Wip'),'Others') AS 'Insuff WIP Type',
PRIORITIZED_REQUESTED_EDC as 'EDC Prioritized Requested Date',
PRIORITIZED_REVISED_EDC as 'EDC Prioritized Revised Date',
if(ecc.check_status in (8,10,11,12),fn_user_name(VQC_REVIEWER_ID),'') as 'VQC done by',
ecm.CLIENT_CASE_EXPECTED_CLOSURE_DATE,
(SELECT EDC_AT_REPORT_SENT
 FROM ec_case_reports ecr WHERE ecm.case_id=ecr.case_id
 AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) EDC_AT_REPORT_SENT

FROM ec_case_master ecm
LEFT JOIN ec_case_fields ecff ON ecm.case_id=ecff.case_id
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_check_queues ecq ON ecc.check_queue=ecq.queue_id AND ecc.check_id=ecq.check_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_user_details eud ON ecc.check_verifier=eud.user_id
LEFT JOIN ec_user_details eud1 ON ecm.documented_by=eud1.user_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_city emcy ON emcy.city_id=eccvs.org_id
LEFT JOIN ec_master_state ems ON emcy.state_id=ems.state_id
LEFT JOIN ec_checks ec ON ecc.check_id=ec.check_id
left join ec_check_families ecf on ec.family_id=ecf.family_id
LEFT JOIN ec_master_disposition emd ON ecc.check_disposition_id=emd.disposition_id
WHERE case_status NOT IN (8,14)
AND check_status <> 9
and ecc.check_id<>193
AND ecm.case_ars_no IN ({in_list})
"""

ANTECEDENT_QUERY = """
SELECT ecc.case_check_id,
	check_name,
	field_name,
	stated_data,
	verified_data,
CLOSURE_COMMENTS
FROM ec_case_checks ecc
LEFT JOIN ec_case_check_data eccd ON ecc.case_check_id=eccd.case_check_id
LEFT JOIN ec_check_fields ecf ON eccd.field_id=ecf.field_id
WHERE ecc.case_check_id IN ({in_list})
"""
