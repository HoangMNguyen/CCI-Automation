#!/usr/bin/env python3
import pandas as pd
import numpy as np
from util import *
from DSMB.DSMB_util import *
from dateutil.relativedelta import *
from datetime import datetime, date
from typing import Optional

def DSMB12423(data, export = False, output_dir = "C:/Users/Hoang Nguyen/Dropbox/Current Work/Download", output_file_name = datetime.now().strftime("%Y%m%d%H%M%S") + "-12423-DSMB Report", debug = False):
    # TODO: DEMO ENROLLMENT LISTING
    if not data['DM'].empty:
        #Subject
        enrollment_df = data['DM'][['Subject']].copy()
        enrollment_df = enrollment_df.sort_values(['Subject'])
        #Cohort Assignment
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DSCA', 'Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)', 'Cohort Assignment')
        # Disease Type
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'NHLMHDIAG', 'Primary Diagnosis of NHL (IG_NS_NA_NHLMHDIAG1.CL_NS_YH_NHLDIAG_cl_NS_NHLDIAG1)', 'Disease NHL')
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'NHLMHDIAG', 'Specify Other Diagnosis (IG_NS_NA_NHLMHDIAG1.TX_NS_NH_NHLDIAGOTH)', 'Disease NHL2')
        enrollment_df['Disease Type'] = None
        # List the columns in the order you want to use them for filling 'Disease'
        columns_to_fill_from = ['Disease NHL', 'Disease NHL2']
        # Use fillna() in a loop to fill 'Disease' from the specified columns
        for col in columns_to_fill_from:
            enrollment_df['Disease Type'] = enrollment_df['Disease Type'].fillna(enrollment_df[col])
        enrollment_df = enrollment_df.drop(columns = columns_to_fill_from)
        #Dose Level
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DSDLA', 'Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)', 'Dose Level')
        # Legal Sex
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)', 'Legal Sex')
        # Sex Assigned at Birth
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)', 'Sex Assigned at Birth')
        # Gender Identity
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Gender Identity (IG_NS_NA_DM1.CL_NS_NH_GENDERID_cl_NS_DMSEX2)', 'Gender Identity')
        #Age at Consent
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)', 'Date of Birth')
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Apheresis Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)', 'Consent Date')
        enrollment_df['Consent Date'] = pd.to_datetime(enrollment_df['Consent Date'])
        enrollment_df['Date of Birth'] = pd.to_datetime(enrollment_df['Date of Birth'])
        mask = ~enrollment_df[['Consent Date', 'Date of Birth']].isnull().any(axis=1)
        enrollment_df.loc[mask, 'Age at Consent'] = enrollment_df[mask].apply(lambda x: relativedelta(x['Consent Date'], x['Date of Birth']).years, axis=1)
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Main Consent Date (IG_NS_NA_IE1.DT_NS_NH_MAINCDAT)', 'Main Consent Date')
        enrollment_df['Main Consent Date'] = pd.to_datetime(enrollment_df['Main Consent Date'])
        # for rows that 'Apheresis Consent Date' isnull but 'Main Consent Date' is not null, then use 'Main Consent Date' instead to calculate age
        enrollment_df.loc[(enrollment_df['Consent Date'].isnull() & enrollment_df['Main Consent Date'].notnull()), 'Age at Consent'] = enrollment_df.loc[(enrollment_df['Consent Date'].isnull() & enrollment_df['Main Consent Date'].notnull())].apply(lambda x: relativedelta(x['Main Consent Date'], x['Date of Birth']).years, axis=1)
        
        #Race
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)', 'Race')
        #Ethnicity
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DM', 'Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)', 'Ethnicity')
        #Subject meets all study eligibility?
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Subject Meets All Study Eligibility (IG_NS_NA_IE3.CL_NS_YH_ELIGYN_cl_YS_YN1)', 'Subject meets all study eligibility?')
        #Reason for Screen Failure
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Other Screen Fail Reason (IG_NS_NA_IE4.TX_NS_YH_OTHRSFREAS)', 'Reason for Screen Failure')
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Screen Failure Reason (IG_NS_NA_IE4.CL_NS_YH_IECAT_cl_NS_IEREASSF1)', 'SF1')
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Select the Primary Inclusion Criterion Excluding this Subject (IG_NS_NA_IE4.CL_NS_NH_ITESTCD_cl_NS_IEINCL1)', 'SF2')
        enrollment_df['SF2'] = enrollment_df[enrollment_df['SF2'].notna()]['SF2'].astype(str)
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'IE', 'Select the Primary Exclusion Criterion Excluding this Subject (IG_NS_NA_IE4.CL_NS_NH_ETESTCD_cl_NS_IEEXCL1)', 'SF3')
        enrollment_df['SF3'] = enrollment_df[enrollment_df['SF3'].notna()]['SF3'].astype(str)
        enrollment_df['SF4'] = enrollment_df['SF1'].fillna('') + " " + enrollment_df['SF2'].fillna('') + enrollment_df['SF3'].fillna('')
        enrollment_df['Reason for Screen Failure'].fillna(enrollment_df['SF4'], inplace=True)
        enrollment_df = enrollment_df.drop(columns = ['SF1', 'SF2', 'SF3', 'SF4'])
        #Infused
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'EXINF', 'Event Group Label', 'Event Group Label')
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'EXINF', 'Was infusion administered? (IG_NS_NA_EXINF1.CL_NS_NH_INFADMIN_cl_YS_YN1)', 'Infused')
        enrollment_df = enrollment_df[enrollment_df['Event Group Label'] != 'Day 0-R']
        enrollment_df = enrollment_df.drop(columns = ['Event Group Label'])
        enrollment_df = add_rename_column_corelisting(enrollment_df, data, 'DSEOS', 'End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)', 'End of Study Date')
        # Update 'Infused' column based on the conditions: 
        enrollment_df.loc[(enrollment_df['Infused'] != 'Yes') & (enrollment_df['End of Study Date'].isnull()), 'Infused'] = 'Pending'
        enrollment_df.loc[(enrollment_df['Infused'] != 'Yes') & (~enrollment_df['End of Study Date'].isnull()), 'Infused'] = 'No'
        enrollment_df = enrollment_df.drop(columns = ['End of Study Date'])
        enrollment_df = enrollment_df.drop_duplicates()
        
        ### TODO: Demo Stats Table
        # !Update this filter options to each cohort
        filter_options = [enrollment_df['Consent Date'].notna() | enrollment_df['Main Consent Date'].notna(), 
                            enrollment_df['Cohort Assignment'] == 'Cohort A: Non-Hodgkin Lymphoma']
        status_list = []
        LegalSex_list = []
        Age_at_Consent_list = []
        Race_list = []
        Ethnicity_list = []
        
        if debug:
            print(len(enrollment_df), enrollment_df.index)
            for filter_option in filter_options:
                print(len(filter_option), filter_option.index)
                print(filter_option.equals(enrollment_df.index))
            
        for filter_index, filter_option in enumerate(filter_options):
            # Apply the filter to the dataframe
            filtered_df = enrollment_df[filter_option].copy()
            filtered_df = filtered_df[(filtered_df['Consent Date'].notna()) | (filtered_df['Main Consent Date'].notna())]
            # Calculate the stats
            ## Total Consented
            TT_df = filtered_df.copy()
            TT = filtered_df['Subject'].count()
            ## Screen Failed
            SF_df = filtered_df[filtered_df['Subject meets all study eligibility?'] == 'No'].copy()
            SF = SF_df['Subject'].count()
            ## Eligible
            EL_df = filtered_df[filtered_df['Subject meets all study eligibility?'] == 'Yes'].copy()
            EL = EL_df['Subject'].count()
            ## Infused
            INF_df = filtered_df[filtered_df['Infused'] == 'Yes'].copy()
            INF = INF_df['Subject'].count()
            
            # Define a dictionary containing the status of each variable
            status_list.append({'Total Consented' : TT, 'Screen Failed' : SF, 'Eligible' : EL, 'Infused' : INF})
            
            # Calculate the stats for the filtered dataframe
            LegalSex_list.append(get_stats_percentage('Legal Sex', TT_df, SF_df, EL_df, INF_df))
            Age_at_Consent_list.append(get_stats_df('Age at Consent', TT_df, SF_df, EL_df, INF_df))
            Race_list.append(get_stats_percentage('Race', TT_df, SF_df, EL_df, INF_df))
            Ethnicity_list.append(get_stats_percentage('Ethnicity', TT_df, SF_df, EL_df, INF_df))
        
        if debug:
            # 0: All Cohorts, 1: Cohort A
            print(status_list)
            print(LegalSex_list)
            print(Age_at_Consent_list)
            print(Race_list)
            print(Ethnicity_list)
        
        # *: remove after calculating the stats
        enrollment_df = enrollment_df.drop(columns = ['Consent Date', 'Date of Birth', 'Main Consent Date'])

        
        ### TODO: INFUSION LISTING
        #adding Target Cell Dose dictionary
        # !: Update this dictionary to the new study
        TCD_dict = {'Dose Level -1 (DL-1)' : 2000000,
                'Dose Level 1 (DL1)' : 7000000,
                'Dose Level 2 (DL2)' : 20000000,
                'Dose Level 3 (DL3)' : 60000000,
                'Not Assigned' : 'Not Assigned'}
        
        # *: PREPARE DATA FOR INFUSION LISTING
        EXCHMO_df = data['EXCHMO'].copy()
        # select unique subject and Event Group Label
        grouped_df = EXCHMO_df.groupby(['Subject', 'Event Group Label'])['Medication (IG_NS_NA_EXCHMO2.CL_NS_NH_EXCCAT_cl_NS_EXCCAT1)'].unique()
        # convert the unique list to string by joining the list with '+' if the list has more than 1 medication
        grouped_df = grouped_df.apply(lambda x: ' + '.join(str(val) for val in x if pd.notna(val)) if len(x) > 1 else x[0]).reset_index()
        # replace the Event Group Label with Day 0 and Day 0-R
        grouped_df.loc[(grouped_df['Event Group Label'] == 'Lymphodepleting Chemotherapy'), 'Event Group Label'] = 'Day 0'
        grouped_df.loc[(grouped_df['Event Group Label'] == 'Retreatment Lymphodepleting Chemotherapy'), 'Event Group Label'] = 'Day 0-R'
        # reassign the dataframe to EXCHMO_df with subject, Study Day, and Medication
        EXCHMO_df = grouped_df

        # TODO: INFUSION LISTING Day 0
        #Subject
        infusion_df = data['DM'][['Subject']].copy()
        infusion_df = infusion_df.sort_values(['Subject'])
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'Event Group Label', 'Event Group Label')
        infusion_df = infusion_df[infusion_df['Event Group Label'] == 'Day 0']
        
        #Cohort Assignment
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'DSCA', 'Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)', 'Cohort Assignment')
        #Dose Level
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'DSDLA', 'Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)', 'Dose Level Assignment')
        #Lymphodepleting Chemotherapy Regimen
        infusion_df = add_rename_column_df(infusion_df, EXCHMO_df[EXCHMO_df['Event Group Label'] == 'Day 0'], 'EXCHMO', 'Medication (IG_NS_NA_EXCHMO2.CL_NS_NH_EXCCAT_cl_NS_EXCCAT1)', 'Lymphodepleting Chemotherapy Regimen')
        # Fill NaN with 
        #Infusion Date
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)', 'Date of TmCD19-IL18 Infusion', 'Subject', 'Event Group Label')
        # convert the date to datetime object and format it to MM-DD-YYYY
        infusion_df['Date of TmCD19-IL18 Infusion'] = infusion_df['Date of TmCD19-IL18 Infusion'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d').strftime('%m-%d-%Y') if pd.notna(x) else x)

        #adding Target Cell Dose using TCD_dict
        infusion_df ['Target Cell Dose'] = infusion_df['Dose Level Assignment'].map(TCD_dict)
        
        #Total huCart19-IL18 Cell Dose
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'CAR T Cell Dose Administered (IG_NS_NA_EXINF1.NM_NS_NH_TDOS)', 'Total TmCD19-IL18 CAR T Cell Dose Administered', 'Subject', 'Event Group Label')
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)', 'Subject', 'Event Group Label')
        #combine Total TmCD19-IL18 CAR T Cell Dose Administered and x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1) columns, compare the new value with 'Target Cell Dose', and convert the Total TmCD19-IL18 CAR T Cell Dose Administered column to string
        infusion_df['Total TmCD19-IL18 CAR T Cell Dose Administered'] = infusion_df['Total TmCD19-IL18 CAR T Cell Dose Administered'].multiply(10**infusion_df['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)'])
        infusion_df = infusion_df.drop(columns = ['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)'])
        
        #Total Cell Dose Administered column
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'Total Cell Dose Administered (IG_NS_NA_EXINF1.NM_NS_NH_TOTDOS)', 'Total Cell Dose Administered', 'Subject', 'Event Group Label')
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)', 'Subject', 'Event Group Label')
        infusion_df['Total Cell Dose Administered'] = infusion_df['Total Cell Dose Administered'].multiply(10**infusion_df['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)'])                
        infusion_df = infusion_df.drop(columns = ['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)'])
        
        # Adding Met Target Dose column based on the condition of Total Cell Dose Administered and Total TmCD19-IL18 CAR T Cell Dose Administered if 'Target Cell Dose' is integer
        infusion_df['Met Target Dose'] = infusion_df.apply(lambda row: "Y" if isinstance(row['Target Cell Dose'], int) and row['Total TmCD19-IL18 CAR T Cell Dose Administered'] >= row['Target Cell Dose'] else "", axis=1)
        infusion_df['Met Target Dose'] = infusion_df.apply(lambda row: "N" if isinstance(row['Target Cell Dose'], int) and row['Total TmCD19-IL18 CAR T Cell Dose Administered'] < row['Target Cell Dose'] else row['Met Target Dose'], axis=1)

        #%scFv Flow
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'Transduction Efficiency (%) (IG_NS_NA_EXINF1.NM_NS_NH_TRANSEFFP)', '%scFv Flow', 'Subject', 'Event Group Label')

        #adding Met Target %scFv
        infusion_df = add_rename_column_corelisting(infusion_df, data, 'EXINF', 'Transduction Efficiency (%) (IG_NS_NA_EXINF1.NM_NS_NH_TRANSEFFP)', 'Met Target %scFv', 'Subject', 'Event Group Label')
        # fillter out the rows that have NaN in Met Target %scFv
        infusion_df['Met Target %scFv'] = infusion_df[infusion_df['Met Target %scFv'].notna()]['Met Target %scFv'].apply(lambda x: "Y" if x >= 2 else "N") 
        # fill NaN with empty string
        infusion_df = infusion_df.fillna("")
        
        # Only keep the rows that have Event Group Label
        infusion_df = infusion_df[infusion_df['Event Group Label'] != '']
        if debug:
            print(infusion_df)

        #TODO: Infusion Listing Day 0-R

        #Subject
        infusionR_df = data['DM'][['Subject']].copy()
        infusionR_df = infusionR_df.sort_values(['Subject'])
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'Event Group Label', 'Event Group Label')
        infusionR_df = infusionR_df[infusionR_df['Event Group Label'] == 'Day 0-R']
        
        #Cohort Assignment
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'DSCA', 'Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)', 'Cohort Assignment')
        #Lymphodepleting Chemotherapy Regimen
        infusionR_df = add_rename_column_df(infusionR_df, EXCHMO_df[EXCHMO_df['Event Group Label'] == 'Day 0-R'], 'EXCHMO', 'Medication (IG_NS_NA_EXCHMO2.CL_NS_NH_EXCCAT_cl_NS_EXCCAT1)', 'Lymphodepleting Chemotherapy Regimen')
        #Infusion Date
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)', 'Date of TmCD19-IL18 Infusion', 'Subject', 'Event Group Label')
        # convert the date to datetime object and format it to MM-DD-YYYY
        infusionR_df['Date of TmCD19-IL18 Infusion'] = infusionR_df['Date of TmCD19-IL18 Infusion'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d').strftime('%m-%d-%Y') if pd.notna(x) else x)
        
        #Total huCart19-IL18 Cell Dose
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'CAR T Cell Dose Administered (IG_NS_NA_EXINF1.NM_NS_NH_TDOS)', 'Total TmCD19-IL18 CAR T Cell Dose Administered', 'Subject', 'Event Group Label')
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)', 'Subject', 'Event Group Label')
        #combine Total TmCD19-IL18 CAR T Cell Dose Administered and x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1) columns, compare the new value with 'Target Cell Dose', and convert the Total TmCD19-IL18 CAR T Cell Dose Administered column to string
        infusionR_df['Total TmCD19-IL18 CAR T Cell Dose Administered'] = infusionR_df['Total TmCD19-IL18 CAR T Cell Dose Administered'].multiply(10**infusionR_df['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)'])
        infusionR_df = infusionR_df.drop(columns = ['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TDOSXP_cl_YS_EX10POW1)'])

        #Total Cell Dose Administered column
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'Total Cell Dose Administered (IG_NS_NA_EXINF1.NM_NS_NH_TOTDOS)', 'Total Cell Dose Administered', 'Subject', 'Event Group Label')
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)', 'x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)', 'Subject', 'Event Group Label')
        infusionR_df['Total Cell Dose Administered'] = infusionR_df['Total Cell Dose Administered'].multiply(10**infusionR_df['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)'])                
        infusionR_df = infusionR_df.drop(columns = ['x 10 to the power of (IG_NS_NA_EXINF1.CL_NS_NH_TOTDOSXP_cl_YS_EX10POW1)'])
        
        #%scFv Flow
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'Transduction Efficiency (%) (IG_NS_NA_EXINF1.NM_NS_NH_TRANSEFFP)', '%scFv Flow', 'Subject', 'Event Group Label')

        #adding Met Target %scFv
        infusionR_df = add_rename_column_corelisting(infusionR_df, data, 'EXINF', 'Transduction Efficiency (%) (IG_NS_NA_EXINF1.NM_NS_NH_TRANSEFFP)', 'Met Target %scFv', 'Subject', 'Event Group Label')
        # fillter out the rows that have NaN in Met Target %scFv
        infusionR_df['Met Target %scFv'] = infusionR_df[infusionR_df['Met Target %scFv'].notna()]['Met Target %scFv'].apply(lambda x: "Y" if x >= 2 else "N") 
        # fill NaN with empty string
        infusionR_df = infusionR_df.fillna("")
        
        # Only keep the rows that have Event Group Label
        infusionR_df = infusionR_df[infusionR_df['Event Group Label'] != '']

        # TODO: INFUSION STATISTICS
        infusion_count = []
        # * Cohort A: Non-Hodgkin Lymphoma
        #Create a new dataframe for Total huCAR T Cell Dose Administered table with infusion_df
        infusionA_df = infusion_df[infusion_df['Cohort Assignment'] == 'Cohort A: Non-Hodgkin Lymphoma']
        infusion_statA1 = get_stats_df('Total TmCD19-IL18 CAR T Cell Dose Administered', infusionA_df)
        #Create a new dataframe for Total Cell Dose Administered table with infusion_df
        infusion_statA2 = get_stats_df('Total Cell Dose Administered', infusionA_df)
        # Count the number of subjects that met the target dose
        met_target_count = infusionA_df[infusionA_df['Met Target Dose'] == 'Y' ].count()['Subject']
        # Count the number of subjects
        total_subject_count = infusionA_df['Subject'].nunique()
        infusion_statA2['Met Target Dose'] = str(met_target_count) + " (" + str(round(met_target_count/total_subject_count*100, 2)) + "%)"
        # Create a new dataframe for %scFv Flow table with infusion_df
        infusion_statA3 = get_stats_perc_df('%scFv Flow', infusionA_df)
        # Count the number of subjects that met the target %scFv
        met_target_count = infusionA_df[infusionA_df['Met Target %scFv'] == 'Y'].count()['Subject']
        infusion_statA3['Met Target %scFv'] = str(met_target_count) + " (" + str(round(met_target_count/total_subject_count*100, 2)) + "%)"
        # Combine the three dataframes
        infusion_statA = pd.concat([infusion_statA1, infusion_statA2, infusion_statA3], axis=1)
        infusion_statA = infusion_statA.replace([np.inf, -np.inf], "")
        infusion_statA = infusion_statA.fillna("")
        infusion_count.append(total_subject_count)
        
        if debug:
            print(infusion_statA)
            print(infusion_count)

        ## TODO: FORMATTING THE DATAFRAME
        # TODO: Day 0
        # Convert the columns to scientific notation if the value is not NaN
        infusion_df['Target Cell Dose'] = infusion_df['Target Cell Dose'].apply(lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x)
        infusion_df['Total TmCD19-IL18 CAR T Cell Dose Administered'] = infusion_df['Total TmCD19-IL18 CAR T Cell Dose Administered'].apply(lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x)
        infusion_df['Total Cell Dose Administered'] = infusion_df['Total Cell Dose Administered'].apply(lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x)
        #adding '%' sign to %scFv Flow
        infusion_df['%scFv Flow'] = infusion_df.apply(lambda row: str(x) + '%' if pd.notna(x := row['%scFv Flow']) else x, axis=1)
        if debug:
            print(infusion_df)
        # TODO: Day 0-R
        # Convert the columns to scientific notation if the value is not NaN
        infusionR_df['Total TmCD19-IL18 CAR T Cell Dose Administered'] = infusionR_df['Total TmCD19-IL18 CAR T Cell Dose Administered'].apply(lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x)
        infusionR_df['Total Cell Dose Administered'] = infusionR_df['Total Cell Dose Administered'].apply(lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x)
        #adding '%' sign to %scFv Flow
        infusionR_df['%scFv Flow'] = infusionR_df.apply(lambda row: str(x) + '%' if pd.notna(x := row['%scFv Flow']) else x, axis=1)

        # TODO: PREPARE
        # Disease Response NHL PET based dictionary 
        DR_NHL_PET_dict = {'Complete Metabolic Response (CMR)' : 1, 'Partial Metabolic Response (PMR)' : 2, 'No Metabolic Response (NMR)' : 3, 'Indeterminate Response (IR)' : 4, 'Progressive Metabolic Disease (PMD)' : 5, 'Not Assessed' : 6, 'Not Reported': 10}
        # Disease Response NHL CT based dictionary
        DR_NHL_CT_dict = {'Complete Radiologic Response (CR)' : 1, 'Partial Response (PR)' : 2, 'Stable Disease (SD)' : 3, 'Indeterminate Response (IR)' : 4, 'Progressive Disease (PD)' : 5, 'Not Assessed' : 6, 'Not Reported': 10}

        # Event Label Update dictionary for cohort A
        event_A_dict = {'Primary Treatment and Follow-up': 'Primary Treatment', 'Primary Retreatment and Follow-up': 'Primary Retreatment', 'Pre-Retreatment Safety Visit' : 'Pre-Retreatment', 'Long Term Follow-up Months 3-60': 'Long Term Follow-up', 'Retreatment Long Term Follow-up Months 3-60': 'Retreatment Long Term Follow-up'}

        # Get data from Initiation of Long Term Follow up
        PD_df = data['DSINITLF'][['Subject', 'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)', 'End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)', 'Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1)']].copy()
        # Filter the data to only subject with 'Disease progression' in Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1) column
        PD_df = PD_df[PD_df['Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1)'] == 'Disease progression']
        # Filter the data to subject in Primary Follow up
        PD_df = PD_df[PD_df['From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)'] == 'Primary Follow-Up']
        # Convert End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT) to datetime object
        PD_df['End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)'] = pd.to_datetime(PD_df['End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)'])
        
        # Get data from Initiation of Long Term Follow up
        PD_Retx_df = data['DSINITLF'][['Subject', 'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)', 'End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)', 'Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1)']].copy()
        # Filter the data to only subject with 'Disease progression' in Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1) column
        PD_Retx_df = PD_Retx_df[PD_Retx_df['Provide reason the Subject is entering into the Long-Term Follow-Up Phase (IG_NS_NA_DSINITLF2.CL_NS_NH_LTFUREAS_cl_NS_LTFURE1)'] == 'Disease progression']
        # Filter the data to subject in Retreatment
        PD_Retx_df = PD_Retx_df[PD_Retx_df['From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)'] == 'Retreatment']
        # Convert End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT) to datetime object
        PD_Retx_df['End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)'] = pd.to_datetime(PD_Retx_df['End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)'])
        
        
        # Get data from DSINITRT
        DSINITRT_df = data['DSINITRT'][['Subject', 'From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)', 'End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)']].copy()
        # Convert End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT) to datetime object
        DSINITRT_df['End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)'] = pd.to_datetime(DSINITRT_df['End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)'])
        
        # TODO: RESPONSE LISTING NHL ONLY
        # Response data dataframe for NHL only
        responseA_df = data['NHLRS'][['Subject', 'Event Group Label', 'Event Date', 'Study Phase (IG_NS_NA_NHLRS1.CL_YS_NH_STUDPSRS_cl_NS_STUDYPS2)', 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)', 'For Unscheduled Primary Treatment Time Point, Specify Day #  (IG_NS_NA_NHLRS1.TX_YS_YH_RSTUDYDAY)',  'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)', 'For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_NHLRS1.TX_NS_YH_RSTUDYDAYR)', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)','CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']].copy()
        responseA_df = responseA_df.sort_values(by=['Subject', 'Event Date'])
        # replace Not Assessed with Not Reported in all columns
        responseA_df = responseA_df.replace('Not Assessed', 'Not Reported')

        # TODO: Cohort A - NHL Primary
        # Filter to only Primary Treatment
        responseA_primary_df = responseA_df[responseA_df['Study Phase (IG_NS_NA_NHLRS1.CL_YS_NH_STUDPSRS_cl_NS_STUDYPS2)'] == 'Primary Treatment']
        # Replace value of "Unscheduled" in column Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
        temp_mask = responseA_primary_df['Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)'] == 'Unscheduled'
        responseA_primary_df = convert_integers_to_strings(responseA_primary_df, 'For Unscheduled Primary Treatment Time Point, Specify Day #  (IG_NS_NA_NHLRS1.TX_YS_YH_RSTUDYDAY)')
        responseA_primary_df.loc[temp_mask, 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)'] = "Day " + responseA_primary_df.loc[temp_mask, 'For Unscheduled Primary Treatment Time Point, Specify Day #  (IG_NS_NA_NHLRS1.TX_YS_YH_RSTUDYDAY)'].fillna(0).astype(int).astype(str)
        # Remove rows with Pre-Treatment Safety Visit
        responseA_primary_df = responseA_primary_df[responseA_primary_df['Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)'] != 'Pre-Treatment Safety Visit' ]
        # Convert Event Date to datetime object
        responseA_primary_df['Event Date'] = pd.to_datetime(responseA_primary_df['Event Date'])
        # Snapshot the responseA_primary_df
        responseA_primary_df_snapshot = responseA_primary_df.copy()
        # check the number of subject for cohort B - CLL Retreatment
        subject_A_prim_count = len(responseA_primary_df['Subject'].unique())
        
        # Check if there is any subject. If yes, then proceed, else skip
        if subject_A_prim_count > 0:
            # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)
            responseA_primary_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].fillna('Not Reported', inplace=True)
            # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)
            responseA_primary_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].fillna('Not Reported', inplace=True)
            # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
            responseA_primary_df['PET-Score'] = responseA_primary_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].map(DR_NHL_PET_dict)
            responseA_primary_df['CT-Score'] = responseA_primary_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].map(DR_NHL_CT_dict)
            if debug:
                print("\nresponseA_primary_df")
                print(responseA_primary_df)
            
            # * CURRENT RESPONSE
            # Filter responseA_primary_df to only subject with PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) not equal to "Not Reported"
            responseA_primary_PET_df = responseA_primary_df[responseA_primary_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] != 'Not Reported']
            # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
            PET_idx = responseA_primary_PET_df.groupby('Subject')['Event Date'].idxmax()
            # Select these rows for the current response
            responseA_primary_current_PET_df = responseA_primary_PET_df.loc[PET_idx].copy()
            # Rename the column 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)' to 'PET Current Time Point'
            responseA_primary_current_PET_df.rename(columns={'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)': 'PET Current Time Point'}, inplace=True)
            # only keep the columns 'Subject', 'PET Current Time Point', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'
            responseA_primary_current_PET_df = responseA_primary_current_PET_df[['Subject', 'PET Current Time Point', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']]
            unique_subjects = pd.DataFrame(responseA_primary_df['Subject'].unique(), columns=['Subject'])
            final_responseA_primary_df = pd.merge(unique_subjects, responseA_primary_current_PET_df, on='Subject', how='left')

            # Filter responseA_primary_df to only subject with CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) not equal to "Not Reported"
            responseA_primary_CT_df = responseA_primary_df[responseA_primary_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] != 'Not Reported']
            # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
            CT_idx = responseA_primary_CT_df.groupby('Subject')['Event Date'].idxmax()
            # Select these rows for the current response
            responseA_primary_current_CT_df = responseA_primary_CT_df.loc[CT_idx].copy()
            # Rename the column 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)' to 'CT Current Time Point'
            responseA_primary_current_CT_df.rename(columns={'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)': 'CT Current Time Point'}, inplace=True)
            # only keep the columns 'Subject', 'CT Current Time Point', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'
            responseA_primary_current_CT_df = responseA_primary_current_CT_df[['Subject', 'CT Current Time Point', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']]

            final_responseA_primary_df = pd.merge(final_responseA_primary_df, responseA_primary_current_CT_df, on='Subject', how='left')
            if debug:
                print("final_responseA_primary_df")
                print(final_responseA_primary_df)


            # * BEST RESPONSE 
            ## Best PET-Based NHL Disease Response primary
            # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
            responseA_best_PET_idx = responseA_primary_df.groupby('Subject')['PET-Score'].idxmin()
            # Select these rows for the best PET-based response
            responseA_best_PET_df = responseA_primary_df.loc[responseA_best_PET_idx].copy()
            # Select the columns subject and PET-Based NHL Disease Response from responseA_best_PET_df
            responseA_best_PET_df = responseA_best_PET_df[['Subject', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)', 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)']]
            # Rename the column PET-Based NHL Disease Response to PET-Based Response
            responseA_best_PET_df.rename(columns={'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)': 'PET-Based Response', 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)' : 'Best PET Time Point'}, inplace=True)
            # Merge left with the primary current response dataframe
            final_responseA_primary_df = pd.merge(final_responseA_primary_df,responseA_best_PET_df, on='Subject', how='left')

            
            ## Best CT-Based NHL Disease Response primary
            # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
            responseA_best_CT_idx = responseA_primary_df.groupby('Subject')['CT-Score'].idxmin()
            # Select these rows for the best CT-based response
            responseA_best_CT_df = responseA_primary_df.loc[responseA_best_CT_idx]
            # Select the columns subject and CT-Based NHL Disease Response from responseA_best_CT_df
            responseA_best_CT_df = responseA_best_CT_df[['Subject', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)', 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)']]
            # Rename the column CT-Based NHL Disease Response to CT-Based Response
            responseA_best_CT_df.rename(columns={'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)': 'CT-Based Response', 'Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)' : 'Best CT Time Point'}, inplace=True)
            # Merge left with the primary response dataframe
            final_responseA_primary_df = pd.merge(final_responseA_primary_df, responseA_best_CT_df, on='Subject', how='left')
            
            if debug:
                print("best response")
                print(final_responseA_primary_df)

            ## Overall NHL Disease Response at Day 28 primary
            # Filter responseA_primary_df to only Day 28
            responseA_primary_D28_df = responseA_df[responseA_df['Primary Treatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPT_cl_NS_RSTPT1)'] == 'Day 28']
            # Selec the columns subject and PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) from responseA_primary_D28_df
            responseA_primary_D28_df = responseA_primary_D28_df[['Subject', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)', 'Event Date']]
            # Compare responseA_primary_D28_df with responseA_df, and add the subjects (do it once) that are not in responseA_primary_D28_df to responseA_primary_D28_df
            responseA_primary_D28_df = pd.concat([responseA_primary_D28_df, responseA_df[~responseA_df['Subject'].isin(responseA_primary_D28_df['Subject'])][['Subject']]])
            # Remove duplicates
            responseA_primary_D28_df = responseA_primary_D28_df.drop_duplicates(subset=['Subject'])
            # Copy snapshot of responseA_primary_df to a temporary dataframe
            temp_df = responseA_primary_df_snapshot
            # Sort the temporary dataframe by Subject and Event Date
            temp_df = temp_df.sort_values(by=['Subject', 'Event Date'])
            # remove all the rows that have nan in PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) and CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)
            temp_df = temp_df[temp_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].notna() | temp_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].notna()]
            # Create a for loop that will check the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) and CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of each subject in responseA_primary_D28_df
            for index, row in responseA_primary_D28_df.iterrows():
                # check if the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the subject is nan, and check if the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the subject is nan
                if pd.isna(row['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']) and pd.isna(row['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']):
                    # if yes, check to see if the subject in in PD_df
                    if row['Subject'] in PD_df['Subject'].values:
                        # get the End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT) date of the subject in PD_df
                        end_date = PD_df[PD_df['Subject'] == row['Subject']]['End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)'].values[0]
                        # find the response of the same subject with the latest event date that is before the Day 28 event date
                        filtered_df = temp_df[(temp_df['Subject'] == row['Subject'])  & (temp_df['Event Date'] <= end_date)]
                        # Check if the filtered DataFrame is empty
                        if not filtered_df.empty:
                            # Access the last row if the DataFrame is not empty
                            temp_row = filtered_df.iloc[-1]
                            # Replace the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the subject with the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the temp_row
                            responseA_primary_D28_df.loc[index, 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] = temp_row['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']
                            # Replace the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the subject with the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the temp_row
                            responseA_primary_D28_df.loc[index, 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] = temp_row['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']
                    # check if the subject is in Initiation of REtx before Day 28
                    elif row['Subject'] in DSINITRT_df['Subject'].values:
                        # get the End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT) date of the subject in DSINITRT_df
                        end_date = DSINITRT_df[DSINITRT_df['Subject'] == row['Subject']]['End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)'].values[0]
                        # find the response of the same subject with the latest event date that is before the Day 28 event date
                        filtered_df = temp_df[(temp_df['Subject'] == row['Subject'])  & (temp_df['Event Date'] <= end_date)]
                        # Check if the filtered DataFrame is empty
                        if not filtered_df.empty:
                            # Access the last row if the DataFrame is not empty
                            temp_row = filtered_df.iloc[-1]
                            # Replace the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the subject with the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the temp_row
                            responseA_primary_D28_df.loc[index, 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] = temp_row['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']
                            # Replace the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the subject with the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the temp_row
                            responseA_primary_D28_df.loc[index, 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] = temp_row['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']
            # Rename the column PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) to PET-Based ORR
            responseA_primary_D28_df.rename(columns={'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)': 'PET-Based ORR', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)': 'CT-Based ORR'}, inplace=True)
            # Merge left with the current response dataframe
            final_responseA_primary_df = pd.merge(final_responseA_primary_df, responseA_primary_D28_df, on='Subject', how='left')
            # Fill NaN with "Not Reported" in column PET-Based ORR
            final_responseA_primary_df['PET-Based ORR'].fillna('Not Reported', inplace=True)
            # Fill NaN with "Not Reported" in column CT-Based ORR
            final_responseA_primary_df['CT-Based ORR'].fillna('Not Reported', inplace=True)

            ## Checking AE and SAE for NHL primary
            # Getting AE and SAE dataframes
            responseA_primary_AE_df = data['AE'][['Subject', 'Form ILB Status', 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)', 'Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)']]
            # Check responseA_primary_AE_df if the subject of responseA_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_primary_df, else add 'N'
            final_responseA_primary_df['AE'] = final_responseA_primary_df['Subject'].apply(lambda x: 'Y' if x in responseA_primary_AE_df['Subject'].values else 'N')
            # Check responseA_primary_AE_df if the subject of responseA_primary_df has SAE in column 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)' . If yes, then add 'Y' to the column 'SAE' in responseA_primary_df, else add 'N'
            final_responseA_primary_df['SAE'] = final_responseA_primary_df['Subject'].apply(lambda x: 'Y' if x in responseA_primary_AE_df[responseA_primary_AE_df['AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)'] == 'SAE']['Subject'].values else 'N')

            ## Checking Study Status for NHL primary
            # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
            responseA_primary_SV_df = data['DSSV'][['Subject', 'Event Label', 'Event Date']]
            # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
            responseA_primary_DSSVLTFU_df = data['DSSVLTFU'][['Subject', 'Event Label', 'Event Date']]
            # Combine DSSVLTFU with SV dataframe vertically
            responseA_primary_SV_df = pd.concat([responseA_primary_SV_df, responseA_primary_DSSVLTFU_df])
            # Sort the dataframe by Subject and Event Date
            responseA_primary_SV_df = responseA_primary_SV_df.sort_values(by=['Subject', 'Event Date'])
            # For each unique subject, get the last row of the dataframe
            responseA_primary_SV_df = responseA_primary_SV_df.groupby('Subject').tail(1)
            # Merge left with the current response dataframe
            final_responseA_primary_df = pd.merge(final_responseA_primary_df, responseA_primary_SV_df[['Subject', 'Event Label']], on='Subject', how='left')
            # Rename the column Event Label to Event Label (Study Status)
            final_responseA_primary_df['Event Label'] = final_responseA_primary_df['Event Label'].map(event_A_dict)

            # Select the columns needed only
            final_responseA_primary_df = final_responseA_primary_df[['Subject',
                                                        'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)',
                                                        'PET Current Time Point',
                                                        'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)',
                                                        'CT Current Time Point',
                                                        'PET-Based Response',
                                                        'Best PET Time Point',
                                                        'CT-Based Response',
                                                        'Best CT Time Point',
                                                        'PET-Based ORR',
                                                        'CT-Based ORR' ,
                                                        'AE',
                                                        'SAE',
                                                        'Event Label']]
            final_responseA_primary_df = final_responseA_primary_df.replace([np.nan, np.inf, -np.inf], '')
            if debug:
                print(final_responseA_primary_df)
            # * Formatting the dataframe

            # TODO: Cohort A - NHL Retreatment
            # Filter to only Primary Treatment
            responseA_retreatment_df = responseA_df[responseA_df['Study Phase (IG_NS_NA_NHLRS1.CL_YS_NH_STUDPSRS_cl_NS_STUDYPS2)'] == 'Retreatment']
            # Replace value of "Unscheduled" in column Retreatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Retreatment  Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
            temp_mask = responseA_retreatment_df['Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)'] == 'Unscheduled'
            # Convert the column For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_NHLRS1.TX_NS_YH_RSTUDYDAYR) to string
            responseA_retreatment_df = convert_integers_to_strings(responseA_retreatment_df, 'For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_NHLRS1.TX_NS_YH_RSTUDYDAYR)')
            responseA_retreatment_df.loc[temp_mask, 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)'] = "Day " + responseA_retreatment_df.loc[temp_mask, 'For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_NHLRS1.TX_NS_YH_RSTUDYDAYR)'].fillna(0).astype(int).astype(str)
            # Remove rows with Pre-Treatment Safety Visit
            responseA_retreatment_df = responseA_retreatment_df[responseA_retreatment_df['Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)'] != 'Pre-Retreatment Safety Visit']
            # Convert Event Date to datetime object
            responseA_retreatment_df['Event Date'] = pd.to_datetime(responseA_retreatment_df['Event Date'])
            # Snapshot the responseA_retreatment_df
            responseA_retreatment_df_snapshot = responseA_retreatment_df.copy()
            # check the number of subject for cohort B - CLL Retreatment
            subject_A_retx_count = len(responseA_retreatment_df['Subject'].unique())
            
            # Check if there is any subject. If yes, then proceed, else skip
            if subject_A_retx_count > 0:
                # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)
                responseA_retreatment_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].fillna('Not Reported', inplace=True)
                # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)
                responseA_retreatment_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].fillna('Not Reported', inplace=True)
                # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
                responseA_retreatment_df['PET-Score'] = responseA_retreatment_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].map(DR_NHL_PET_dict)
                responseA_retreatment_df['CT-Score'] = responseA_retreatment_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].map(DR_NHL_CT_dict)
                
                # * CURRENT RESPONSE
                # Filter responseA_retreatment_df to only subject with PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) not equal to "Not Reported"
                responseA_retreatment_PET_df = responseA_retreatment_df[responseA_retreatment_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] != 'Not Reported']
                # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
                PET_idx = responseA_retreatment_PET_df.groupby('Subject')['Event Date'].idxmax()
                # Select these rows for the current response
                responseA_retreatment_current_PET_df = responseA_retreatment_PET_df.loc[PET_idx].copy()
                # Rename the column 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)' to 'PET Current Time Point'
                responseA_retreatment_current_PET_df.rename(columns={'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)': 'PET Current Time Point'}, inplace=True)
                # only keep the columns 'Subject', 'PET Current Time Point', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'
                responseA_retreatment_current_PET_df = responseA_retreatment_current_PET_df[['Subject', 'PET Current Time Point', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']]
                if debug:
                    print(responseA_retreatment_current_PET_df)

                # Filter responseA_retreatment_df to only subject with CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) not equal to "Not Reported"
                responseA_retreatment_CT_df = responseA_retreatment_df[responseA_retreatment_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] != 'Not Reported']
                # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
                CT_idx = responseA_retreatment_CT_df.groupby('Subject')['Event Date'].idxmax()
                # Select these rows for the current response
                responseA_retreatment_current_CT_df = responseA_retreatment_CT_df.loc[CT_idx].copy()
                # Rename the column 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)' to 'CT Current Time Point'
                responseA_retreatment_current_CT_df.rename(columns={'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)': 'CT Current Time Point'}, inplace=True)
                # only keep the columns 'Subject', 'CT Current Time Point', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'
                responseA_retreatment_current_CT_df = responseA_retreatment_current_CT_df[['Subject', 'CT Current Time Point', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']]
                if debug:
                    print(responseA_retreatment_current_CT_df)
                    
                final_responseA_retreatment_df = pd.merge(responseA_retreatment_current_PET_df, responseA_retreatment_current_CT_df, on='Subject', how='left')

                # * BEST RESPONSE 
                ## Best PET-Based NHL Disease Response primary
                # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
                responseA_best_PET_idx = responseA_retreatment_df.groupby('Subject')['PET-Score'].idxmin()
                # Select these rows for the best PET-based response
                responseA_best_PET_df = responseA_retreatment_df.loc[responseA_best_PET_idx].copy()
                # Select the columns subject and PET-Based NHL Disease Response from responseA_best_PET_df
                responseA_best_PET_df = responseA_best_PET_df[['Subject', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)', 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)']]
                # Rename the column PET-Based NHL Disease Response to PET-Based Response
                responseA_best_PET_df.rename(columns={'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)': 'PET-Based Response', 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)' : 'Best PET Time Point'}, inplace=True)
                # Fill NaN with "Not Reported" in column PET-Based Response
                responseA_best_PET_df['PET-Based Response'].fillna('Not Reported', inplace=True)
                # Merge left with the primary current response dataframe
                final_responseA_retreatment_df = pd.merge(final_responseA_retreatment_df,responseA_best_PET_df, on='Subject', how='left')

                
                ## Best CT-Based NHL Disease Response primary
                # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
                responseA_best_CT_idx = responseA_retreatment_df.groupby('Subject')['CT-Score'].idxmin()
                # Select these rows for the best CT-based response
                responseA_best_CT_df = responseA_retreatment_df.loc[responseA_best_CT_idx]
                # Select the columns subject and CT-Based NHL Disease Response from responseA_best_CT_df
                responseA_best_CT_df = responseA_best_CT_df[['Subject', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)', 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)']]
                # Rename the column CT-Based NHL Disease Response to CT-Based Response
                responseA_best_CT_df.rename(columns={'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)': 'CT-Based Response', 'Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)' : 'Best CT Time Point'}, inplace=True)
                # Fill NaN with "Not Reported" in column CT-Based Response
                responseA_best_CT_df['CT-Based Response'].fillna('Not Reported', inplace=True)
                # Merge left with the primary response dataframe
                final_responseA_retreatment_df = pd.merge(final_responseA_retreatment_df, responseA_best_CT_df, on='Subject', how='left')
                

                ## * Overall NHL Disease Response at Day 28-R
                # Filter responseA_retreatment_df to only Day 28-R
                responseA_retreatment_D28_df = responseA_df[responseA_df['Retreatment Time Point (IG_NS_NA_NHLRS1.CL_NS_NH_RSTPTR_cl_NS_RSTPT2)'] == 'Day 28-R']
                # Selec the columns subject and PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) from responseA_retreatment_D28_df
                responseA_retreatment_D28_df = responseA_retreatment_D28_df[['Subject', 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)', 'Event Date']]
                # Compare responseA_retreatment_D28_df with responseA_df, and add the subjects (do it once) that are not in responseA_retreatment_D28_df to responseA_retreatment_D28_df
                responseA_retreatment_D28_df = pd.concat([responseA_retreatment_D28_df, responseA_df[~responseA_df['Subject'].isin(responseA_retreatment_D28_df['Subject'])][['Subject']]])
                # Remove duplicates
                responseA_retreatment_D28_df = responseA_retreatment_D28_df.drop_duplicates(subset=['Subject'])
                
                # Copy snapshot of responseA_retreatment_df to a temporary dataframe
                temp_df = responseA_retreatment_df_snapshot
                # Sort the temporary dataframe by Subject and Event Date
                temp_df = temp_df.sort_values(by=['Subject', 'Event Date'])
                # remove all the rows that have nan in PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) and CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)
                temp_df = temp_df[temp_df['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'].notna() | temp_df['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'].notna()]
                # Create a for loop that will check the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) and CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of each subject in responseA_retreatment_D28_df
                for index, row in responseA_retreatment_D28_df.iterrows():
                    # check if the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the subject is nan, and check if the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the subject is nan
                    if pd.isna(row['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']) and pd.isna(row['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']):
                        # if yes, check to see if the subject in in PD_Retx_df
                        if row['Subject'] in PD_Retx_df['Subject'].values:
                            # get the End of Retreatment Date (ig_INITLF1.DSENRETXDAT) date of the subject in PD_Retx_df
                            end_date = PD_Retx_df[PD_Retx_df['Subject'] == row['Subject']]['End of Retreatment Date (ig_INITLF1.DSENRETXDAT)'].values[0]
                            # if yes, find the response of the same subject with the latest event date that is before the Day 28 event date
                            filtered_df = temp_df[(temp_df['Subject'] == row['Subject'])  & (temp_df['Event Date'] <= end_date)]
                            # Check if the filtered DataFrame is empty
                            if not filtered_df.empty:
                                # Access the last row if the DataFrame is not empty
                                temp_row = filtered_df.iloc[-1]
                                # Replace the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the subject with the PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) of the temp_row
                                responseA_retreatment_D28_df.loc[index, 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] = temp_row['PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)']
                                # Replace the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the subject with the CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1) of the temp_row
                                responseA_retreatment_D28_df.loc[index, 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] = temp_row['CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)']
                # Rename the column PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1) to PET-Based ORR
                responseA_retreatment_D28_df.rename(columns={'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)': 'PET-Based ORR', 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)': 'CT-Based ORR'}, inplace=True)
                # Merge left with the current response dataframe
                final_responseA_retreatment_df = pd.merge(final_responseA_retreatment_df, responseA_retreatment_D28_df, on='Subject', how='left')
                # Fill NaN with "Not Reported" in column PET-Based ORR
                final_responseA_retreatment_df['PET-Based ORR'].fillna('Not Reported', inplace=True)
                # Fill NaN with "Not Reported" in column CT-Based ORR
                final_responseA_retreatment_df['CT-Based ORR'].fillna('Not Reported', inplace=True)
                
                ## Checking AE and SAE for NHL primary
                # Getting AE and SAE dataframes
                responseA_retreatment_AE_df = data['AE'][['Subject', 'Form ILB Status', 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)', 'Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)']]
                # Check responseA_retreatment_AE_df if the subject of responseA_retreatment_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_retreatment_df, else add 'N'
                final_responseA_retreatment_df['AE'] = final_responseA_retreatment_df['Subject'].apply(lambda x: 'Y' if x in responseA_retreatment_AE_df['Subject'].values else 'N')
                # Check responseA_retreatment_AE_df if the subject of responseA_retreatment_df has SAE in column 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)' . If yes, then add 'Y' to the column 'SAE' in responseA_retreatment_df, else add 'N'
                final_responseA_retreatment_df['SAE'] = final_responseA_retreatment_df['Subject'].apply(lambda x: 'Y' if x in responseA_retreatment_AE_df[responseA_retreatment_AE_df['AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)'] == 'SAE']['Subject'].values else 'N')

                ## Checking Study Status for NHL primary
                # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
                responseA_retreatment_SV_df = data['DSSV'][['Subject', 'Event Label', 'Event Date']]
                # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
                responseA_retreatment_DSSVLTFU_df = data['DSSVLTFU'][['Subject', 'Event Label', 'Event Date']]
                # Combine DSSVLTFU with SV dataframe vertically
                responseA_retreatment_SV_df = pd.concat([responseA_retreatment_SV_df, responseA_retreatment_DSSVLTFU_df])
                # Sort the dataframe by Subject and Event Date
                responseA_retreatment_SV_df = responseA_retreatment_SV_df.sort_values(by=['Subject', 'Event Date'])
                # For each unique subject, get the last row of the dataframe
                responseA_retreatment_SV_df = responseA_retreatment_SV_df.groupby('Subject').tail(1)
                # Merge left with the current response dataframe
                final_responseA_retreatment_df = pd.merge(final_responseA_retreatment_df, responseA_retreatment_SV_df[['Subject', 'Event Label']], on='Subject', how='left')
                
                # * Formatting the dataframe
                # Rename the column Event Label to Event Label (Study Status)
                final_responseA_retreatment_df['Event Label'] = final_responseA_retreatment_df['Event Label'].map(event_A_dict)
                
                # Select the columns needed only
                final_responseA_retreatment_df = final_responseA_retreatment_df[['Subject',
                                                            'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)',
                                                            'PET Current Time Point',
                                                            'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)',
                                                            'CT Current Time Point',
                                                            'PET-Based Response',
                                                            'Best PET Time Point',
                                                            'CT-Based Response',
                                                            'Best CT Time Point',
                                                            'PET-Based ORR',
                                                            'CT-Based ORR' ,
                                                            'AE',
                                                            'SAE',
                                                            'Event Label']]
                final_responseA_retreatment_df = final_responseA_retreatment_df.replace([np.nan, np.inf, -np.inf], '')
                if debug:
                    print(final_responseA_retreatment_df)
                
        ### TODO: REPONSE STATS
        # TODO: SAFETY STATS
        
        # Gather all stats of Cohort A
        total_infused_df = infusion_df.copy()
        # Getting AE and SAE dataframes
        AE_df = data['AE'][['Subject', 'Form ILB Status', 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)', 'Event Onset (IG_NS_NA_AE1.CL_NS_YH_AEONSET_cl_NS_AEONSET1)']].copy()
        # Check responseA_primary_AE_df if the subject of responseA_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_primary_df, else add 'N'
        total_infused_df['AE'] = total_infused_df['Subject'].apply(lambda x: 'Y' if x in AE_df['Subject'].values else 'N')
        # Check responseA_primary_AE_df if the subject of responseA_primary_df has SAE in column 'AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)' . If yes, then add 'Y' to the column 'SAE' in responseA_primary_df, else add 'N'
        total_infused_df['SAE'] = total_infused_df['Subject'].apply(lambda x: 'Y' if x in AE_df[AE_df['AE or SAE? (IG_NS_NA_AE2.CL_YS_YH_AESEV_cl_NS_AESAE1)'] == 'SAE']['Subject'].values else 'N')
        
        # # Total number of subjects in cohort A, B, and C
        AE_total_count = get_stats_percentage('AE', total_infused_df).T
        SAE_total_count = get_stats_percentage('SAE', total_infused_df).T
        # merge AE and SAE dataframes
        safety_total_df = pd.concat([AE_total_count, SAE_total_count], axis=1)
        
        # TODO: RESPONSE STATS
        if subject_A_prim_count > 0:
            responseA_stat = final_responseA_primary_df.copy()
            # replace 'Not Assessed' with 'Not Reported' for all columns in responseA_stat
            responseA_stat = responseA_stat.replace('Not Assessed', 'Not Reported')
            response_stat_A_BOR_PET = get_stats_percentage('PET-Based Response', responseA_stat)
            response_stat_A_BOR_CT = get_stats_percentage('CT-Based Response', responseA_stat)
            response_stat_A_ORR_PET = get_stats_percentage('PET-Based ORR', responseA_stat)
            response_stat_A_ORR_CT = get_stats_percentage('CT-Based ORR', responseA_stat)

        # TODO: UPDATE FORMAT for cohort A after getting the stats

        if subject_A_prim_count > 0:
            final_responseA_primary_df.loc[(final_responseA_primary_df['Event Label'] == 'Retreatment Long Term Follow-up') | (final_responseA_primary_df['Event Label'] == 'Primary Retreatment') | (final_responseA_primary_df['Event Label'] == 'Pre-Retreatment'), 'PET-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLPET_cl_NS_RSNHLPET1)'] = 'Transitioned to Retreatment'
            final_responseA_primary_df.loc[(final_responseA_primary_df['Event Label'] == 'Retreatment Long Term Follow-up') | (final_responseA_primary_df['Event Label'] == 'Primary Retreatment') | (final_responseA_primary_df['Event Label'] == 'Pre-Retreatment'), 'CT-Based NHL Disease Response (IG_NS_NA_NHLRS2.CL_NS_NH_NHLCT_cl_NS_RSNHLCT1)'] = 'Transitioned to Retreatment'
            final_responseA_primary_df.loc[(final_responseA_primary_df['Event Label'] == 'Retreatment Long Term Follow-up') | (final_responseA_primary_df['Event Label'] == 'Primary Retreatment') | (final_responseA_primary_df['Event Label'] == 'Pre-Retreatment'), 'PET Current Time Point'] = 'Transitioned to Retreatment'
            final_responseA_primary_df.loc[(final_responseA_primary_df['Event Label'] == 'Retreatment Long Term Follow-up') | (final_responseA_primary_df['Event Label'] == 'Primary Retreatment') | (final_responseA_primary_df['Event Label'] == 'Pre-Retreatment'), 'CT Current Time Point'] = 'Transitioned to Retreatment'

            
    if export:
        with pd.ExcelWriter(output_dir  + '/' + output_file_name + '.xlsx', engine = 'xlsxwriter') as writer:  
            # TODO: - Add formatting and coloring
            # TODO: - for each tab: write data, format data, write header, format header
            
            ## * FORMATING AND COLORING
            bold_11_format = writer.book.add_format({'bg_color': '#FFFFFF',
                                                        'text_wrap': True,
                                                        'valign': 'vcenter',
                                                        'align': 'center',
                                                        'bold' : True,
                                                        'font_name' : 'Calibri',
                                                        'font_size' : 11,
                                                        'border': 1})
            bold_12_format = writer.book.add_format({'bg_color': '#FFFFFF',
                                                        'text_wrap': False,
                                                        'valign': 'vcenter',
                                                        'align': 'center',
                                                        'bold' : True,
                                                        'font_name' : 'Calibri',
                                                        'font_size' : 12,
                                                        'border': 1})
            bold_12_wrap_format = writer.book.add_format({'bg_color': '#FFFFFF',
                                                        'text_wrap': True,
                                                        'valign': 'vcenter',
                                                        'align': 'center',
                                                        'bold' : True,
                                                        'font_name' : 'Calibri',
                                                        'font_size' : 12,
                                                        'border': 1})
            bold_11_wrap_format = writer.book.add_format({'bg_color': '#FFFFFF',
                                                        'text_wrap': True,
                                                        'valign': 'vcenter',
                                                        'align': 'center',
                                                        'bold' : True,
                                                        'font_name' : 'Calibri',
                                                        'font_size' : 11,
                                                        'border': 1})
            normal_data_format = writer.book.add_format({'bg_color': '#FFFFFF',
                                                        'text_wrap': False,
                                                        'valign': 'vcenter',
                                                        'align': 'center',
                                                        'bold' : False,
                                                        'font_name' : 'Calibri',
                                                        'font_size' : 11,
                                                        'border': 1})
            # Create a format for a black cell
            black_cell = writer.book.add_format({'bg_color': 'black'})
            if data['DM']['Subject'].count() > 0:
                ## TODO: DSMB-Demo Stats Table
                if enrollment_df['Subject'].count() > 0:
                    # * WRITING DATA: LegalSex_list, Age_at_Consent_list, Race_list, Ethnicity_list
                    worksheet1 = writer.book.add_worksheet('DSMB-Demo Stats Table')
                    
                    # * FORMAT DATA 
                    for i in range(0, len(status_list)):
                        for j in range(0, len(LegalSex_list[i])):
                            for k in range(0, len(LegalSex_list[i].columns)):
                                worksheet1.write(j + 3 , k + 1 + i * 4, LegalSex_list[i].iloc[j, k], normal_data_format)
                        for j in range(0, len(Age_at_Consent_list[i])):
                            for k in range(0, len(Age_at_Consent_list[i].columns)):
                                worksheet1.write(j + 8, k + 1 + i * 4, Age_at_Consent_list[i].iloc[j, k], normal_data_format)
                        for j in range(0, len(Race_list[i])):
                            for k in range(0, len(Race_list[i].columns)):
                                worksheet1.write(j + 12, k + 1 + i * 4, Race_list[i].iloc[j, k], normal_data_format)
                        for j in range(0, len(Ethnicity_list[i])):
                            for k in range(0, len(Ethnicity_list[i].columns)):
                                worksheet1.write(j + 22, k + 1 + i * 4, Ethnicity_list[i].iloc[j, k], normal_data_format)
                    
                    
                    # Apply the format to a range of cells
                    # worksheet1.set_column('B:I', None, normal_data_format)

                    # * WRITING HEADER AND FORMATTING
                    Sex_order = ['Male', 'Female', 'Nonbinary (X)', 'Not Reported']
                    Age_order = ['Mean SD', 'Median', 'Range']
                    Race_order = ['African American', 'Alaska Native', 'American Indian', 'Asian','Caucasian', 'Multiple Races', 'Pacific Islander', 'Other', 'Unknown']
                    Ethnicity_order = ['Hispanic', 'Non-Hispanic', 'Unknown']
                    
                    for i in range(0, len(Sex_order)):
                        worksheet1.write(i + 3, 0, Sex_order[i], bold_11_format)
                    for i in range(0, len(Age_order)):
                        worksheet1.write(i + 8, 0, Age_order[i], bold_11_format)
                    for i in range(0, len(Race_order)):
                        worksheet1.write(i + 12, 0, Race_order[i], bold_11_format)
                    for i in range(0, len(Ethnicity_order)):
                        worksheet1.write(i + 22, 0, Ethnicity_order[i], bold_11_format)
                    
                    worksheet1.merge_range('B1:E1', 'Overall Study Enrollment', bold_12_format)
                    worksheet1.merge_range('F1:I1', 'Cohort A (NHL) Enrollment' , bold_12_format)
                    worksheet1.write(1, 0, 'Status', bold_11_format)
                    for i in range(len(status_list)):
                        worksheet1.write(1, 1 + i * 4, "Total Consented\nN=" + str(status_list[i]['Total Consented']), bold_11_wrap_format)
                        worksheet1.write(1, 2+ i * 4, "Screen Failed\nN=" + str(status_list[i]['Screen Failed']), bold_11_wrap_format)
                        worksheet1.write(1, 3+ i * 4, "Eligible\nN=" + str(status_list[i]['Eligible']), bold_11_wrap_format)
                        worksheet1.write(1, 4+ i * 4, "Infused\nN=" + str(status_list[i]['Infused']), bold_11_wrap_format)
                    
                    worksheet1.merge_range('A3:I3', 'Legal Sex' , bold_11_format)
                    worksheet1.merge_range('A8:I8', 'Age at Consent' , bold_11_format)
                    worksheet1.merge_range('A12:I12', 'Race' , bold_11_format)
                    worksheet1.merge_range('A22:I22', 'Ethnicity' , bold_11_format)
                    worksheet1.autofit()
                    
                    ## TODO: Enrollment Listing
                    # * WRITING DATA: enrollment_df
                    worksheet2 = writer.book.add_worksheet('DSMB-Enrollment Listing')
                    # * WRITING HEADER AND FORMATTING
                    # Assuming 'enrollment_df' is your DataFrame
                    enrollment_df.replace([np.inf, -np.inf], np.nan, inplace=True) # Replace INF with NaN
                    enrollment_df.fillna('', inplace=True) # Replace NaN with a placeholder
                    for i in range(0, len(enrollment_df.columns)):
                        worksheet2.write(0, i, enrollment_df.columns[i], bold_11_format)
                    # * FORMAT DATA 
                    for i in range(0, len(enrollment_df)):
                        for j in range(0, len(enrollment_df.columns)):
                            worksheet2.write(i + 1, j, enrollment_df.iloc[i, j], normal_data_format)
                    # Autofit
                    worksheet2.autofit()
                    
                    ## TODO: DSMB-New Infusion Statistics
                    # * WRITING DATA: new_infusion_df
                    # * WRITING DATA: enrollment_df
                    worksheet3 = writer.book.add_worksheet('DSMB-New Infusion Statistics')
                    
                    # * FORMATING DATA
                    for i in range(0, len(infusion_statA)):
                        for j in range(0, len(infusion_statA.columns)):
                            worksheet3.write(i + 3, j + 1, infusion_statA.iloc[i, j], normal_data_format)

                    # * WRITING HEADER AND FORMATTING
                    stat_order = ['Mean SD', 'Median', 'Range']

                    worksheet3.merge_range('B1:D1', 'Cells Infused' , bold_12_wrap_format)
                    worksheet3.merge_range('E1:F1', 'Transduction Efficiency' , bold_12_wrap_format)
                    worksheet3.write('B2', 'Total Cells' , bold_12_wrap_format)
                    worksheet3.write('C2', 'TmCD19-IL18 Cells' , bold_12_wrap_format)
                    worksheet3.write('D2', 'Met Target Dose' , bold_12_wrap_format)
                    worksheet3.write('E2', '%scFv Flow' , bold_12_wrap_format)
                    worksheet3.write('F2', 'Met Target %scFv' , bold_12_wrap_format)
                    worksheet3.merge_range('A3:F3', 'Cohort A (N=' + str(infusion_count[0]) + ')' , bold_12_wrap_format)
                    
                    # Merge and format data
                    worksheet3.merge_range('D4:D6', infusion_statA.iloc[0, 2] , normal_data_format)
                    worksheet3.merge_range('F4:F6', infusion_statA.iloc[0, 4] , normal_data_format)

                    for i in range(0, len(stat_order)):
                        worksheet3.write(i + 3, 0, stat_order[i], bold_11_format)

                    # * Autofit
                    worksheet3.autofit()

                    ## TODO: DSMB-Infusion Listing
                    worksheet4 = writer.book.add_worksheet('DSMB-Infusion Listing')
                    # * WRITING AND FORMATING DATA
                    for i in range(0, len(infusion_df)):
                        for j in range(0, len(infusion_df.columns)):
                            worksheet4.write(i + 2, j, infusion_df.iloc[i, j], normal_data_format)
                    for i in range(0, len(infusionR_df)):
                        for j in range(0, len(infusionR_df.columns)):
                            worksheet4.write(i + 2, j + 14, infusionR_df.iloc[i, j], normal_data_format)
                    # * WRITING HEADER AND FORMATTING
                    worksheet4.merge_range('A1:A2', 'Subject ID' , bold_12_wrap_format)
                    worksheet4.merge_range('B1:B2', 'Study Day (Primary)' , bold_12_wrap_format)
                    worksheet4.merge_range('C1:C2', 'Cohort Assignment' , bold_12_wrap_format)
                    worksheet4.merge_range('D1:D2', 'Dose Level Assignment' , bold_12_wrap_format)
                    worksheet4.merge_range('E1:E2', 'Lymphodepleting Chemotherapy Regimen' , bold_12_wrap_format)
                    worksheet4.merge_range('F1:F2', 'Date of TmCD19-IL18 Infusion' , bold_12_wrap_format)
                    worksheet4.merge_range('G1:J1', 'Cells Infused' , bold_12_wrap_format)
                    worksheet4.merge_range('K1:L1', 'Transduction Efficiency' , bold_12_wrap_format)
                    worksheet4.write('G2', 'Target Cell Dose' , bold_12_wrap_format)
                    worksheet4.write('H2', 'Total TmCD19-IL18 CAR T Cell Dose Administered' , bold_12_wrap_format)
                    worksheet4.write('I2', 'Total Cell Dose Administered' , bold_12_wrap_format)
                    worksheet4.write('J2', 'Met Target Dose' , bold_12_wrap_format)
                    worksheet4.write('K2', '%scFv Flow' , bold_12_wrap_format)
                    worksheet4.write('L2', 'Met Target %scFv' , bold_12_wrap_format)

                    worksheet4.merge_range('O1:O2', 'Subject ID' , bold_12_wrap_format)
                    worksheet4.merge_range('P1:P2', 'Study Day (Retreatment)' , bold_12_wrap_format)
                    worksheet4.merge_range('Q1:Q2', 'Cohort Assignment' , bold_12_wrap_format)
                    worksheet4.merge_range('R1:R2', 'Lymphodepleting Chemotherapy Regimen' , bold_12_wrap_format)
                    worksheet4.merge_range('S1:S2', 'Date of TmCD19-IL18 Retreatment Infusion' , bold_12_wrap_format)
                    worksheet4.merge_range('T1:U1', 'Cells Infused' , bold_12_wrap_format)
                    worksheet4.merge_range('V1:W1', 'Transduction Efficiency' , bold_12_wrap_format)
                    worksheet4.write('T2', 'Total TmCD19-IL18 CAR T Cell Dose Administered' , bold_12_wrap_format)
                    worksheet4.write('U2', 'Total Cell Dose Administered' , bold_12_wrap_format)
                    worksheet4.write('V2', '%scFv Flow' , bold_12_wrap_format)
                    worksheet4.write('W2', 'Met Target %scFv' , bold_12_wrap_format)

                    # Autofit
                    worksheet4.autofit()
                    
                    ## TODO: DSMB-Response Stats
                    worksheet5 = writer.book.add_worksheet('DSMB-Response Stats')
                    # * WRITING DATA
                    # * FORMATING DATA
                    # Safety Data
                    for i in range(0, len(safety_total_df)):
                        for j in range(0, len(safety_total_df.columns)):
                            worksheet5.write(i + 3, j, safety_total_df.iloc[i, j], normal_data_format)
                    # Response Data Cohort A
                    if subject_A_prim_count == 0:
                        for i in range(0, 6):
                            worksheet5.write(i + 8, 1, '0 (0%)', normal_data_format)
                            worksheet5.write(i + 8, 3, '0 (0%)', normal_data_format)
                            worksheet5.write(i + 15, 1, '0 (0%)', normal_data_format)
                            worksheet5.write(i + 15, 3, '0 (0%)', normal_data_format)
                    else:
                        for i in range(0, len(response_stat_A_BOR_PET)):
                            for j in range(0, len(response_stat_A_BOR_PET.columns)):
                                worksheet5.write(i + 8, j + 1, response_stat_A_BOR_PET.iloc[i, j], normal_data_format)
                        for i in range(0, len(response_stat_A_BOR_CT)):
                            for j in range(0, len(response_stat_A_BOR_CT.columns)):
                                worksheet5.write(i + 8, j + 3, response_stat_A_BOR_CT.iloc[i, j], normal_data_format)
                        for i in range(0, len(response_stat_A_ORR_PET)):
                            for j in range(0, len(response_stat_A_ORR_PET.columns)):
                                worksheet5.write(i + 15, j + 1, response_stat_A_ORR_PET.iloc[i, j], normal_data_format)
                        for i in range(0, len(response_stat_A_ORR_CT)):
                            for j in range(0, len(response_stat_A_ORR_CT.columns)):
                                worksheet5.write(i + 15, j + 3, response_stat_A_ORR_CT.iloc[i, j], normal_data_format)
                                
                    # * WRITING HEADER AND FORMATTING
                    # Safety Headers
                    # number of subject of safety_total_df
                    safety_total_df_subject_count = len(infusion_df['Subject'].unique())
                    worksheet5.merge_range('A1:D1', 'Safety Statistics (N=' + str(safety_total_df_subject_count) + ')', bold_12_wrap_format)
                    worksheet5.merge_range('A2:B2', 'Adverse Events', bold_11_format)
                    worksheet5.merge_range('C2:D2', 'Serious Adverse Events ', bold_11_format)
                    worksheet5.write('A3', 'Yes', bold_11_format)
                    worksheet5.write('B3', 'No', bold_11_format)
                    worksheet5.write('C3', 'Yes', bold_11_format)
                    worksheet5.write('D3', 'No', bold_11_format)

                    # Response Headers
                    worksheet5.merge_range('A6:D6', 'Cohort A (NHL) Subject Response (N=' + str(subject_A_prim_count) + ')', bold_12_format)
                    worksheet5.merge_range('A7:B7', 'PET-Based Response', bold_11_format)
                    worksheet5.merge_range('C7:D7', 'CT-Based Response', bold_11_format)
                    worksheet5.merge_range('A8:D8', 'Best Overall Response (BOR)', bold_11_format)
                    worksheet5.merge_range('A15:D15', 'Overall Response Rate (ORR) at Day 28', bold_11_format)
                    # Listing Response Criteria
                    response_A_PET = ['Complete Metabolic Response (CMR)', 'Partial Metabolic Response (PMR)', 'No Metabolic Response (NMR)', 'Indeterminate Response (IR)', 'Progressive Metabolic Disease (PMD)', 'Not Reported' ]
                    response_A_CT = ['Complete Radiologic Response (CR)', 'Partial Response (PR)', 'Stable Disease (SD)', 'Indeterminate Response (IR)', 'Progressive Disease (PD)', 'Not Reported' ]
                    for i in range(0, len(response_A_PET)):
                        worksheet5.write(i + 8, 0, response_A_PET[i], bold_11_format)
                        worksheet5.write(i + 15, 0, response_A_PET[i], bold_11_format)
                    for i in range(0, len(response_A_CT)):
                        worksheet5.write(i + 8, 2, response_A_CT[i], bold_11_format)
                        worksheet5.write(i + 15, 2, response_A_CT[i], bold_11_format)
                    worksheet5.autofit()
                    
                    ## TODO: Response Listing NHL
                    if subject_A_prim_count > 0:
                        worksheet6 = writer.book.add_worksheet('Response Listing NHL')
                        # * WRITING DATA
                        # * FORMATING DATA
                        if subject_A_prim_count > 0:
                            for i in range(0, len(final_responseA_primary_df)):
                                for j in range(0, len(final_responseA_primary_df.columns)):
                                    worksheet6.write(i + 3, j , final_responseA_primary_df.iloc[i, j], normal_data_format)
                            if subject_A_retx_count > 0:
                                for i in range(0, len(final_responseA_retreatment_df)):
                                    for j in range(0, len(final_responseA_retreatment_df.columns)):
                                        worksheet6.write(i + 3, j + 16, final_responseA_retreatment_df.iloc[i, j], normal_data_format)
                        # * WRITING HEADER AND FORMATTING
                        if subject_A_prim_count > 0:
                            worksheet6.merge_range('A1:N1', 'Cohort A (NHL)- Primary Follow-up', bold_12_format)
                            worksheet6.merge_range('A2:A3', 'Subject ID', bold_11_format)
                            worksheet6.merge_range('B2:E2', 'Current Response', bold_11_format)
                            worksheet6.merge_range('F2:I2', 'Best Response/Timepoint', bold_11_format)
                            worksheet6.merge_range('J2:K2', 'Overall Response/Day 28', bold_11_format)
                            worksheet6.write('B3', 'PET-Based Response', bold_11_format)
                            worksheet6.write('C3', 'Study Timepoint', bold_11_format)
                            worksheet6.write('D3', 'CT-Based Response', bold_11_format)
                            worksheet6.write('E3', 'Study Timepoint', bold_11_format)
                            worksheet6.write('F3', 'PET-Based Response', bold_11_format)
                            worksheet6.write('G3', 'Study Timepoint', bold_11_format)
                            worksheet6.write('H3', 'CT-Based Response', bold_11_format)
                            worksheet6.write('I3', 'Study Timepoint', bold_11_format)
                            worksheet6.write('J3', 'PET-Based ORR', bold_11_format)
                            worksheet6.write('K3', 'CT-Based ORR', bold_11_format)
                            worksheet6.merge_range('L2:L3', 'Adverse Events \n(Y/N)', bold_11_wrap_format)
                            worksheet6.merge_range('M2:M3', 'Serious Adverse Events \n(Y/N)', bold_11_wrap_format)
                            worksheet6.merge_range('N2:N3', 'Study Status', bold_11_wrap_format)
                            if subject_A_retx_count > 0:
                                worksheet6.merge_range('Q1:AD1', 'Cohort A (NHL)- Retreatment Follow-up', bold_12_format)
                                worksheet6.merge_range('Q2:Q3', 'Subject ID', bold_11_format)
                                worksheet6.merge_range('R2:U2', 'Current Response', bold_11_format)
                                worksheet6.merge_range('V2:Y2', 'Best Response/Timepoint', bold_11_format)
                                worksheet6.merge_range('Z2:AA2', 'Overall Response/Day 28', bold_11_format)
                                worksheet6.write('R3', 'PET-Based Response', bold_11_format)
                                worksheet6.write('S3', 'Study Timepoint', bold_11_format)
                                worksheet6.write('T3', 'CT-Based Response', bold_11_format)
                                worksheet6.write('U3', 'Study Timepoint', bold_11_format)
                                worksheet6.write('V3', 'PET-Based Response', bold_11_format)
                                worksheet6.write('W3', 'Study Timepoint', bold_11_format)
                                worksheet6.write('X3', 'CT-Based Response', bold_11_format)
                                worksheet6.write('Y3', 'Study Timepoint', bold_11_format)
                                worksheet6.write('Z3', 'PET-Based ORR', bold_11_format)
                                worksheet6.write('AA3', 'CT-Based ORR', bold_11_format)
                                worksheet6.merge_range('AB2:AB3', 'Adverse Events \n(Y/N)', bold_11_wrap_format)
                                worksheet6.merge_range('AC2:AC3', 'Serious Adverse Events \n(Y/N)', bold_11_wrap_format)
                                worksheet6.merge_range('AD2:AD3', 'Study Status', bold_11_wrap_format)
                        # Autofit
                        worksheet6.autofit()
