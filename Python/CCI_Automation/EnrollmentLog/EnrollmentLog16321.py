#!/usr/bin/env python3
import pandas as pd
from dateutil.relativedelta import *
from util import *

def EnrollmentLog16321(final_data):         
    # DM
    # if 'DM' in final_data:
    DM_df = final_data['DM'][['Subject','Race (ig_DM1.RACE)', 'Ethnicity (ig_DM1.ETHNIC)', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)', 'Legal Sex (ig_DM1.SEX)', 'Date of Birth (ig_DM1.BRTHDAT)', 'Apheresis Consent Date (ig_DM1.RFICDAT)' ]].copy()
    DM_new_col_name = {'Race (ig_DM1.RACE)': 'Race', 'Ethnicity (ig_DM1.ETHNIC)': 'Ethnicity', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)': 'Sex Assigned at Birth', 'Legal Sex (ig_DM1.SEX)': 'Legal Sex', 'Date of Birth (ig_DM1.BRTHDAT)': 'Date of Birth', 'Apheresis Consent Date (ig_DM1.RFICDAT)' : 'Apheresis Consent Date'}
    DM_df = DM_df.rename(columns=DM_new_col_name)
    sorted_DM_df = DM_df.sort_values(['Subject'])

    # DSCAS

    DSCA_df = final_data['DSCA'][['Subject','Cohort Assignment (ig_DSCA1.CACHASCOD)' ]].copy()
    DSCA_new_col_name = {'Cohort Assignment (ig_DSCA1.CACHASCOD)':'Assigned Cohort'}
    DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
    merged_df = pd.merge(sorted_DM_df, DSCA_df, on='Subject', how='left')
    index_reference = merged_df.columns.get_loc('Race')
    merged_df.insert(index_reference, 'Assigned Cohort', merged_df.pop('Assigned Cohort'))
    # print(merged_df)


    # IE
    IE_df = final_data['IE'][['Subject',  'Main Consent Date (ig_IE1.MAINCDAT)', 'Date of Eligibility Confirmation by Physician-Investigator (ig_IE5.ELIGPIDAT)', 'Date of Completion of Monitoring Visit for Eligibility (ig_IE5.ELIGMONDAT)']].copy()
    IE_new_col_name = {'Main Consent Date (ig_IE1.MAINCDAT)':'Main Consent Date', 'Date of Eligibility Confirmation by Physician-Investigator (ig_IE5.ELIGPIDAT)': 'Date Physician-Investigator Confirmed Eligibility', 'Date of Completion of Monitoring Visit for Eligibility (ig_IE5.ELIGMONDAT)': 'Date of Monitoring Visit for Eligibility'}
    IE_df = IE_df.rename(columns=IE_new_col_name)
    merged_df = pd.merge(merged_df, IE_df, on='Subject', how='left')
    #convert date columsn to datetime type
    merged_df['Apheresis Consent Date'] = pd.to_datetime(merged_df['Apheresis Consent Date'])
    merged_df['Date of Birth'] = pd.to_datetime(merged_df['Date of Birth'])
    # Calculate time difference in days and convert to years
    # create a mask for non-NaT values in the two columns
    mask = ~merged_df[['Apheresis Consent Date', 'Date of Birth']].isnull().any(axis=1)

    # apply relativedelta only to rows with non-NaT values in both columns
    merged_df.loc[mask, 'Age at Consent'] = merged_df[mask].apply(lambda x: relativedelta(x['Apheresis Consent Date'], x['Date of Birth']).years, axis=1)

    merged_df['Apheresis Consent Date'] = pd.to_datetime(merged_df['Apheresis Consent Date']).dt.strftime('%m/%d/%Y')
    merged_df['Main Consent Date'] = pd.to_datetime(merged_df['Main Consent Date']).dt.strftime('%m/%d/%Y')
    merged_df['Date Physician-Investigator Confirmed Eligibility'] = pd.to_datetime(merged_df['Date Physician-Investigator Confirmed Eligibility']).dt.strftime('%m/%d/%Y')
    merged_df['Date of Monitoring Visit for Eligibility'] = pd.to_datetime(merged_df['Date of Monitoring Visit for Eligibility']).dt.strftime('%m/%d/%Y')
    merged_df = merged_df.drop('Date of Birth', axis = 1)
    #move the column
    index_reference = merged_df.columns.get_loc('Apheresis Consent Date')
    merged_df.insert(index_reference, 'Age at Consent', merged_df.pop('Age at Consent'))


    #PRAPH
    APH_df = final_data['PRAPH'][['Subject','Apheresis Type (ig_PRAPH1.APHCAT)', 'Apheresis Date (ig_PRAPH1.APHDAT)']].copy()
    APH_new_col_name = {'Apheresis Type (ig_PRAPH1.APHCAT)': 'Apheresis Type (Fresh or Historical)', 'Apheresis Date (ig_PRAPH1.APHDAT)': 'Date of Apheresis Collection'}
    APH_df = APH_df.rename(columns=APH_new_col_name)
    merged_df = pd.merge(merged_df, APH_df, on='Subject', how='left')
    merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')
    # print(merged_df)

    #EXINF
    EXINF_df = final_data['EXINF'][['Subject', 'Event Group Label', 'Study Treatment Date (ig_EXINF1.INFDAT)']].copy()
    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0']
    EXINF_new_col_name = {'Study Treatment Date (ig_EXINF1.INFDAT)': 'CART T cell Administration Date (Day 0)'}
    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
    merged_df['CART T cell Administration Date (Day 0)'] = pd.to_datetime(merged_df['CART T cell Administration Date (Day 0)']).dt.strftime('%m/%d/%Y')
    # print(merged_df)
    

    #DSINITLF
    INITLF_df = final_data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)']].copy()
    INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] != 'Retreatment']
    INITLF_new_col_name = {'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)': 'Last Study Visit Completed in Primary Follow-Up', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)': 'Initiation of LTFU Date'}
    INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
    INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
    merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
    merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')
    # print(merged_df)

    #DSINITRT
    DSINITRT_df = final_data['DSINITRT'][['Subject','Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', 'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)', 'End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)']].copy()
    DSINITRT_new_col_name = {'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)': 'Retreatment?', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)': 'Phase', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)': 'Initiation of Retx Date'}
    DSINITRT_df = DSINITRT_df.rename(columns=DSINITRT_new_col_name)
    DSINITRT_df['Initiation of Retx Date'] = DSINITRT_df['Initiation of Retx Date'].fillna(DSINITRT_df['End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)'])
    DSINITRT_df = DSINITRT_df.drop('End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)', axis = 1)
    merged_df = pd.merge(merged_df, DSINITRT_df, on='Subject', how='left')
    # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
    merged_df.loc[merged_df['Phase'] == 'Primary Follow-Up', 'Initiation of LTFU Date'] = 'N/A'
    merged_df['Last Study Visit Completed in Primary Follow-Up'] = merged_df['Last Study Visit Completed in Primary Follow-Up'].fillna(merged_df['Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)'])
    merged_df = merged_df.drop(['Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', 'Phase'], axis = 1)
    merged_df['Initiation of Retx Date'] = pd.to_datetime(merged_df['Initiation of Retx Date']).dt.strftime('%m/%d/%Y')

    # EXINF Retx
    EXINF_df = final_data['EXINF'][['Subject','Event Group Label', 'Study Treatment Date (ig_EXINF1.INFDAT)']].copy()
    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0-R1']
    EXINF_new_col_name = {'Study Treatment Date (ig_EXINF1.INFDAT)': 'CAR T Cell Retreatment Date (Day 0-R1)'}
    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
    merged_df['CAR T Cell Retreatment Date (Day 0-R1)'] = pd.to_datetime(merged_df['CAR T Cell Retreatment Date (Day 0-R1)']).dt.strftime('%m/%d/%Y')
    # print(merged_df)

    #INITLF Retx
    INITLF_df = final_data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)', 'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)']].copy()
    INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] == 'Retreatment']
    INITLF_df = INITLF_df[INITLF_df['Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)'] == 'Retreatment-R1']
    INITLF_new_col_name = {'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)': 'Last Study Visit Completed in Retreatment (-R1)', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)': 'Initiation of Retreatment LTFU Date'}
    INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
    INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
    INITLF_df = INITLF_df.drop('Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)', axis = 1)
    merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
    merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')
    
    # EOS
    EOS_df = final_data['DSEOS'][['Subject','End of Study Date (ig_DSEOS1.EOSDAT)' ]].copy()
    EOS_new_col_name = {'End of Study Date (ig_DSEOS1.EOSDAT)':'End of Study Date'}
    EOS_df = EOS_df.rename(columns=EOS_new_col_name)
    merged_df = pd.merge(merged_df, EOS_df, on='Subject', how='left')
    merged_df['End of Study Date'] = pd.to_datetime(merged_df['End of Study Date']).dt.strftime('%m/%d/%Y')
    #update headers and fill N/A
    merged_df = merged_df.rename(columns={'Subject': 'Subject ID#'})
    merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
    # Remove duplicate rows based on both columns
    merged_df.drop_duplicates(inplace=True)
    return merged_df

