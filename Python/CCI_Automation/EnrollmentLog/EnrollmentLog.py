#!/usr/bin/env python3
import pandas as pd
from dateutil.relativedelta import *
from util import *


class EnrollmentLog():
    def __init__(self, study_name, input_dir, output_dir, output_file_name, cut_off_date = None):   #self.selected_option2, self.input_folder_name, self.output_folder_name, self.output_file_name 
        if input_dir == None:
            print("No dir selected!")
            return
        else:
            self.study_name = study_name
            self.input_dir = input_dir
            self.output_dir = output_dir
            self.output_file_name = output_file_name
            self.data = read_data_dict_zip_corelisting(self.input_dir, cut_off_date)
            # self.read_data_dict()
            self.error_message = False
            self.output_df = self.collect_data()
            self.output()
            

    def get_study_name(self):
        folder_name = self.input_dir.split("/")[-1]
        study_name = folder_name.split("_")[2]
        if not study_name.isdigit():
            print("It's not a study name: %s".format(study_name) )
        else:
            return study_name

    def collect_data(self):

            if self.study_name =='03821':
                # DM
                # if 'DM' in self.data:
                DM_df = self.data['DM'][['Subject','Race (ig_DM1.RACE)', 'Ethnicity (ig_DM1.ETHNIC)', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)', 'Legal Sex (ig_DM1.SEX)', 'Date of Birth (ig_DM1.BRTHDAT)', 'Apheresis Consent Date (ig_DM1.RFICDAT)' ]].copy()
                DM_new_col_name = {'Race (ig_DM1.RACE)': 'Race', 'Ethnicity (ig_DM1.ETHNIC)': 'Ethnicity', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)': 'Sex Assigned at Birth', 'Legal Sex (ig_DM1.SEX)': 'Legal Sex', 'Date of Birth (ig_DM1.BRTHDAT)': 'Date of Birth', 'Apheresis Consent Date (ig_DM1.RFICDAT)' : 'Apheresis Consent Date'}
                DM_df = DM_df.rename(columns=DM_new_col_name)
                sorted_DM_df = DM_df.sort_values(['Subject'])

                # DSCA

                DSCA_df = self.data['DSCA'][['Subject','Cohort Assignment (ig_DSCA1.CACHASCOD)' ]].copy()
                DSCA_new_col_name = {'Cohort Assignment (ig_DSCA1.CACHASCOD)':'Assigned Cohort'}
                DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
                merged_df = pd.merge(sorted_DM_df, DSCA_df, on='Subject', how='left')
                index_reference = merged_df.columns.get_loc('Race')
                merged_df.insert(index_reference, 'Assigned Cohort', merged_df.pop('Assigned Cohort'))
                # print(merged_df)


                # MHDIAG
                MHDIAG_df = self.data['MHDIAG'][['Subject','Disease Type (ig_MHDIAG1.RSCAT)' ]].copy()
                MHDIAG_new_col_name = {'Disease Type (ig_MHDIAG1.RSCAT)':'Disease Type'}
                MHDIAG_df = MHDIAG_df.rename(columns=MHDIAG_new_col_name)
                merged_df = pd.merge(merged_df, MHDIAG_df, on='Subject', how='left')
                index_reference = merged_df.columns.get_loc('Assigned Cohort')
                merged_df.insert(index_reference, 'Disease Type', merged_df.pop('Disease Type'))


                # IE
                IE_df = self.data['IE'][['Subject',  'Main Consent Date (ig_IE1.MAINCDAT)', 'Date of eligibility confirmation by physician-investigator (ig_IE5.ELIGPIDAT)', 'Date of completion of monitoring visit for eligibility (ig_IE5.ELIGMONDAT)']].copy()
                IE_new_col_name = {'Main Consent Date (ig_IE1.MAINCDAT)':'Main Consent Date', 'Date of eligibility confirmation by physician-investigator (ig_IE5.ELIGPIDAT)': 'Date Physician-Investigator Confirmed Eligibility', 'Date of completion of monitoring visit for eligibility (ig_IE5.ELIGMONDAT)': 'Date of Monitoring Visit for Eligibility'}
                IE_df = IE_df.rename(columns=IE_new_col_name)
                merged_df = pd.merge(merged_df, IE_df, on='Subject', how='left')
                #convert date columsn to datetime type
                merged_df['Apheresis Consent Date'] = pd.to_datetime(merged_df['Apheresis Consent Date'])
                merged_df['Main Consent Date'] = pd.to_datetime(merged_df['Main Consent Date'])
                merged_df['Date of Birth'] = pd.to_datetime(merged_df['Date of Birth'])
                # Calculate time difference in days and convert to years
                # create a mask for non-NaT values in the two columns
                mask = ~merged_df[['Apheresis Consent Date', 'Date of Birth']].isnull().any(axis=1)

                # apply relativedelta only to rows with non-NaT values in both columns
                merged_df.loc[mask, 'Age at Consent'] = merged_df[mask].apply(lambda x: relativedelta(x['Apheresis Consent Date'], x['Date of Birth']).years, axis=1)
                # for rows that 'Apheresis Consent Date' isnull but 'Main Consent Date' is not null, then use 'Main Consent Date' instead to calculate age
                merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull()), 'Age at Consent'] = merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull())].apply(lambda x: relativedelta(x['Main Consent Date'], x['Date of Birth']).years, axis=1)
                merged_df['Apheresis Consent Date'] = pd.to_datetime(merged_df['Apheresis Consent Date']).dt.strftime('%m/%d/%Y')
                merged_df['Main Consent Date'] = pd.to_datetime(merged_df['Main Consent Date']).dt.strftime('%m/%d/%Y')
                merged_df['Date Physician-Investigator Confirmed Eligibility'] = pd.to_datetime(merged_df['Date Physician-Investigator Confirmed Eligibility']).dt.strftime('%m/%d/%Y')
                merged_df['Date of Monitoring Visit for Eligibility'] = pd.to_datetime(merged_df['Date of Monitoring Visit for Eligibility']).dt.strftime('%m/%d/%Y')
                # print(merged_df)
                merged_df = merged_df.drop('Date of Birth', axis = 1)

                #move the column
                index_reference = merged_df.columns.get_loc('Apheresis Consent Date')
                merged_df.insert(index_reference, 'Age at Consent', merged_df.pop('Age at Consent'))

                #PRAPH
                APH_df = self.data['PRAPH'][['Subject','Apheresis Type (ig_PRAPH1.APHCAT)', 'Apheresis Date (ig_PRAPH1.APHDAT)']].copy()
                APH_new_col_name = {'Apheresis Type (ig_PRAPH1.APHCAT)': 'Apheresis Type (Fresh or Historical)', 'Apheresis Date (ig_PRAPH1.APHDAT)': 'Date of Apheresis Collection'}
                APH_df = APH_df.rename(columns=APH_new_col_name)
                merged_df = pd.merge(merged_df, APH_df, on='Subject', how='left')
                merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')
                

                #EXVCNINF
                EXVCNINF_df = self.data['EXVCNINF'][['Subject', 'Infusion Date (ig_EXVCNINF1.INFDAT)']].copy()
                EXVCNINF_new_col_name = {'Infusion Date (ig_EXVCNINF1.INFDAT)': 'Date of VCN-01 Infusion'}
                EXVCNINF_df = EXVCNINF_df.rename(columns=EXVCNINF_new_col_name)
                merged_df = pd.merge(merged_df, EXVCNINF_df, on='Subject', how='left')
                merged_df['Date of VCN-01 Infusion'] = pd.to_datetime(merged_df['Date of VCN-01 Infusion']).dt.strftime('%m/%d/%Y')
                # print(merged_df)
                

                #EXMESOINF
                INF_df = self.data['EXMESOINF'][['Subject','Event Group Label', 'Infusion Date (ig_EXMESOINF1.INFDAT)']].copy()
                INF_df = INF_df[INF_df['Event Group Label'] != 'Day 0-R']
                INF_new_col_name = {'Infusion Date (ig_EXMESOINF1.INFDAT)': 'Date of huCART-meso Infusion'}
                INF_df = INF_df.rename(columns=INF_new_col_name)
                INF_df = INF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, INF_df, on='Subject', how='left')
                merged_df['Date of huCART-meso Infusion'] = pd.to_datetime(merged_df['Date of huCART-meso Infusion']).dt.strftime('%m/%d/%Y')

                #DSINITLF
                INITLF_df = self.data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] != 'Retreatment']
                INITLF_new_col_name = {'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)': 'Last Study Visit Completed in Primary Follow-Up', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)': 'Initiation of LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')

                #DSINITRT
                DSINITRT_df = self.data['DSINITRT'][['Subject','Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', 'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)','From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)', 'End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)']].copy()
                DSINITRT_new_col_name = {'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)': 'Retreatment?', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)': 'Phase', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)': 'Initiation of Retx Date', 'End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)': 'End of Long-Term Follow-Up Date'}
                DSINITRT_df = DSINITRT_df.rename(columns=DSINITRT_new_col_name)
                merged_df = pd.merge(merged_df, DSINITRT_df, on='Subject', how='left')
                # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
                merged_df.loc[merged_df['Phase'] == 'Primary Follow-Up', 'Initiation of LTFU Date'] = 'N/A'
                merged_df['Last Study Visit Completed in Primary Follow-Up'] = merged_df['Last Study Visit Completed in Primary Follow-Up'].fillna(merged_df['Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)'])
                merged_df['Initiation of Retx Date'] = merged_df['Initiation of Retx Date'].fillna(merged_df['End of Long-Term Follow-Up Date'])
                merged_df = merged_df.drop('Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', axis = 1)
                # Drop End of Long-Term Follow-Up Date and Phase column
                merged_df = merged_df.drop(['End of Long-Term Follow-Up Date','Phase'], axis = 1)
                merged_df['Initiation of Retx Date'] = pd.to_datetime(merged_df['Initiation of Retx Date']).dt.strftime('%m/%d/%Y')

                # INF Retx
                INF_df = self.data['EXMESOINF'][['Subject','Event Group Label', 'Infusion Date (ig_EXMESOINF1.INFDAT)']].copy()
                INF_df = INF_df[INF_df['Event Group Label'] == 'Day 0-R']
                INF_new_col_name = {'Infusion Date (ig_EXMESOINF1.INFDAT)': 'CAR T cell Retreatment Date [Day 0-R]'}
                INF_df = INF_df.rename(columns=INF_new_col_name)
                INF_df = INF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, INF_df, on='Subject', how='left')
                merged_df['CAR T cell Retreatment Date [Day 0-R]'] = pd.to_datetime(merged_df['CAR T cell Retreatment Date [Day 0-R]']).dt.strftime('%m/%d/%Y')


                #INITLF Retx
                INITLF_df = self.data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] == 'Retreatment']
                INITLF_new_col_name = {'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)': 'Last Study Visit Completed in Retreatment F/up', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)': 'Initiation of Retreatment LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')
                
                # EOS
                EOS_df = self.data['DSEOS'][['Subject','End of Study Date (ig_DSEOS1.EOSDAT)' ]].copy()
                EOS_new_col_name = {'End of Study Date (ig_DSEOS1.EOSDAT)':'End of Study Date'}
                EOS_df = EOS_df.rename(columns=EOS_new_col_name)
                merged_df = pd.merge(merged_df, EOS_df, on='Subject', how='left')
                merged_df['End of Study Date'] = pd.to_datetime(merged_df['End of Study Date']).dt.strftime('%m/%d/%Y')
                #update headers and fill N/A
                merged_df = merged_df.rename(columns={'Subject': 'Subject ID#'})
                merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
                
                return merged_df

            if self.study_name =='11823':
                # DM
                # Check missing 
                if 'DM' in self.data:
                    DM_df = self.data['DM'][['Subject','Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)', 'Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)', 'Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)', 'Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)', 'Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)', 'Apheresis Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)' ]].copy()
                    DM_new_col_name = {'Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)': 'Race', 'Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)': 'Ethnicity', 'Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)': 'Sex Assigned at Birth' , 'Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)': 'Legal Sex', 'Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)': 'Date of Birth', 'Apheresis Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)' : 'Apheresis Consent Date'}
                    DM_df = DM_df.rename(columns=DM_new_col_name)
                    DM_df['Apheresis Consent Date'] = pd.to_datetime(DM_df['Apheresis Consent Date'])
                    sorted_DM_df = DM_df.sort_values(['Subject'])
                    merged_df = sorted_DM_df
                    

                #DSDLA
                if 'DSDLA' in self.data:
                    DSDLA_df = self.data['DSDLA'][['Subject','Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)' ]].copy()
                    DSDLA_new_col_name = {'Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)':'Assigned Dose Level'}
                    DSDLA_df = DSDLA_df.rename(columns=DSDLA_new_col_name)
                    merged_df = pd.merge(merged_df, DSDLA_df, on='Subject', how='left')
                    index_reference = merged_df.columns.get_loc('Race')
                    merged_df.insert(index_reference, 'Assigned Dose Level', merged_df.pop('Assigned Dose Level'))

                # IE
                if 'IE' in self.data:
                    # Subset and rename columns simultaneously
                    IE_cols_to_select = {
                        'Subject': 'Subject',
                        'Main Consent Date (IG_NS_NA_IE1.DT_NS_NH_MAINCDAT)': 'Main Consent Date',
                        'Subject Meets All Study Eligibility (IG_NS_NA_IE3.CL_NS_YH_ELIGYN_cl_YS_YN1)': 'Subject Meets All Study Eligibility',
                        'Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)': 'Date Physician-Investigator Confirmed Eligibility',
                        'Date of Completion of Monitoring Visit for Eligibility (IG_NS_NA_IE5.DT_NS_YH_ELIGMONDAT)': 'Date of Monitoring Visit for Eligibility'
                    }
                    IE_df = self.data['IE'][list(IE_cols_to_select.keys())].rename(columns=IE_cols_to_select)
                    # Filter Subject Meets All Study Eligibility == Yes
                    IE_df = IE_df[IE_df['Subject Meets All Study Eligibility'] == 'Yes'].drop('Subject Meets All Study Eligibility', axis=1)
                    #convert date columsn to datetime type
                    IE_df['Date Physician-Investigator Confirmed Eligibility'] = pd.to_datetime(IE_df['Date Physician-Investigator Confirmed Eligibility']).dt.strftime('%m/%d/%Y')
                    IE_df['Date of Monitoring Visit for Eligibility'] = pd.to_datetime(IE_df['Date of Monitoring Visit for Eligibility']).dt.strftime('%m/%d/%Y')
                    merged_df = pd.merge(merged_df, IE_df, on='Subject', how='left')
                    merged_df['Main Consent Date'] = pd.to_datetime(merged_df['Main Consent Date'])
                    merged_df['Date of Birth'] = pd.to_datetime(merged_df['Date of Birth'])
                    
                    # find mask for non-NaT values in the two columns
                    mask = ~merged_df[['Apheresis Consent Date', 'Date of Birth']].isnull().any(axis=1)
                    # apply relativedelta only to rows with non-NaT values in both columns
                    merged_df.loc[mask, 'Age at Consent'] = merged_df[mask].apply(lambda x: relativedelta(x['Apheresis Consent Date'], x['Date of Birth']).years, axis=1)
                    # for rows that 'Apheresis Consent Date' isnull but 'Main Consent Date' is not null, then use 'Main Consent Date' instead to calculate age
                    merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull()), 'Age at Consent'] = merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull())].apply(lambda x: relativedelta(x['Main Consent Date'], x['Date of Birth']).years, axis=1)
                    
                    
                    merged_df['Apheresis Consent Date'] = merged_df['Apheresis Consent Date'].dt.strftime('%m/%d/%Y')
                    merged_df['Main Consent Date'] = merged_df['Main Consent Date'].dt.strftime('%m/%d/%Y')
                    merged_df = merged_df.drop('Date of Birth', axis = 1)
                    #move the column
                    index_reference = merged_df.columns.get_loc('Apheresis Consent Date')
                    merged_df.insert(index_reference, 'Age at Consent', merged_df.pop('Age at Consent'))

                #APH
                if 'PRAPH' in self.data:
                    PRAPH_cols_to_select = {
                        'Subject': 'Subject',
                        'Event Group Label': 'Event Group Label',
                        'Apheresis Type (IG_NS_NA_PRAPH1.CL_NS_YH_APHTP_cl_NS_APHTP1)': 'Apheresis Type (Fresh or Historical)',
                        'Apheresis Date (IG_NS_NA_PRAPH1.DT_NS_NH_APHDAT)': 'Date of Apheresis Collection'
                    }
                    PRAPH_df = self.data['PRAPH'][list(PRAPH_cols_to_select.keys())].rename(columns=PRAPH_cols_to_select)

                    # Filter rows and drop unnecessary column
                    PRAPH_df = PRAPH_df[PRAPH_df['Event Group Label'] == 'Initial Study Enrollment/Apheresis'].drop('Event Group Label', axis=1)
                    merged_df = pd.merge(merged_df, PRAPH_df, on='Subject', how='left')
                    merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')

                #EXCHMO
                if 'EXCHMO' in self.data:
                    EXCHMO_df = self.data['EXCHMO'][['Subject','Event Group Label', 'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)']].copy()
                    EXCHMO_df = EXCHMO_df[EXCHMO_df['Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)'] != 'NaN']
                    EXCHMO_df = EXCHMO_df[EXCHMO_df['Event Group Label'] == 'Lymphodepleting Chemotherapy']
                    EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject'])
                    EXCHMO_new_col_name = {'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)': 'Date of Initiation of LD Chemo'}
                    EXCHMO_df = EXCHMO_df.rename(columns=EXCHMO_new_col_name)
                    EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                    EXCHMO_df = EXCHMO_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                    merged_df['Date of Initiation of LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of LD Chemo']).dt.strftime('%m/%d/%Y')

                #INF
                if 'EXINF' in self.data:
                    EXINF_df = self.data['EXINF'][['Subject','Event Group Label', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)']].copy()
                    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] != 'Day 0-R']
                    EXINF_new_col_name = {'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)': 'CAR T cell Infusion Date [Day 0]'}
                    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                    merged_df['CAR T cell Infusion Date [Day 0]'] = pd.to_datetime(merged_df['CAR T cell Infusion Date [Day 0]']).dt.strftime('%m/%d/%Y')

                #DSINITLF
                if 'DSINITLF' in self.data:
                    DSINITLF_cols_to_select = {
                        'Subject': 'Subject',
                        'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)': 'Phase Entering LTFU',
                        'Last Study Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCPFU_cl_YS_LVCPFU1)': 'Last Study Visit Completed in Primary Follow-Up',
                        'End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)': 'Initiation of LTFU Date'
                    }
                    DSINITLF_df = self.data['DSINITLF'][list(DSINITLF_cols_to_select.keys())].rename(columns=DSINITLF_cols_to_select)

                    # Filter rows and drop unnecessary column
                    DSINITLF_df = DSINITLF_df[DSINITLF_df['Phase Entering LTFU'] == 'Primary Follow-Up'].drop('Phase Entering LTFU', axis=1)
                    # Merge dataframes
                    merged_df = pd.merge(merged_df, DSINITLF_df, on='Subject', how='left')
                    # Format date
                    merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')
                    # print(merged_df)

                #DSINITRT
                if 'DSINITRT' in self.data:
                    DSINITRT_cols_to_select = {
                        'Subject': 'Subject',
                        'Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_YS_LVCPFU1)': 'Last Visit Completed in PFU',
                        'From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)' : 'Phase',
                        'End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)': 'Initiation of Retx Date',
                        'End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)': 'End of Long-Term Follow-Up Date'
                    }
                    DSINITRT_df = self.data['DSINITRT'][list(DSINITRT_cols_to_select.keys())].rename(columns=DSINITRT_cols_to_select)
                    merged_df = pd.merge(merged_df, DSINITRT_df, on='Subject', how='left')
                    # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
                    merged_df.loc[merged_df['Phase'] == 'Primary Follow-Up', 'Initiation of LTFU Date'] = 'N/A'
                    merged_df['Last Study Visit Completed in Primary Follow-Up'] = merged_df['Last Study Visit Completed in Primary Follow-Up'].fillna(merged_df['Last Visit Completed in PFU'])
                    merged_df = merged_df.drop('Last Visit Completed in PFU', axis=1)
                    merged_df['Initiation of Retx Date'] = merged_df['Initiation of Retx Date'].fillna(merged_df['End of Long-Term Follow-Up Date'])
                    # drop End of Long-Term Follow-Up Date column
                    merged_df = merged_df.drop(['End of Long-Term Follow-Up Date', 'Phase'], axis=1)
                    merged_df['Initiation of Retx Date'] = pd.to_datetime(merged_df['Initiation of Retx Date']).dt.strftime('%m/%d/%Y')

                #EXCHMO Retreatment
                if 'EXCHMO' in self.data:
                    # Subset and rename columns simultaneously
                    EXCHMO_cols_to_select = {
                        'Subject': 'Subject',
                        'Event Group Label': 'Event Group Label',
                        'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)': 'Date of Initiation of Retx LD Chemo'
                    }
                    EXCHMO_df = self.data['EXCHMO'][list(EXCHMO_cols_to_select.keys())].rename(columns=EXCHMO_cols_to_select)

                    # Filter rows
                    EXCHMO_df = EXCHMO_df[
                        (EXCHMO_df['Date of Initiation of Retx LD Chemo'] != 'NaN') & 
                        (EXCHMO_df['Event Group Label'] == 'Retreatment Lymphodepleting Chemotherapy')
                    ]

                    # Drop duplicates and unnecessary columns
                    EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject']).drop('Event Group Label', axis=1)
                    # Sort by Subject
                    EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                    merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                    merged_df['Date of Initiation of Retx LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of Retx LD Chemo']).dt.strftime('%m/%d/%Y')
                

                # EXINF Retx
                if 'EXINF' in self.data:
                    EXINF_df = self.data['EXINF'][['Subject','Event Group Label', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)']].copy()
                    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0-R']
                    EXINF_new_col_name = {'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)': 'CAR T cell Infusion Date [Day 0-R]'}
                    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                    merged_df['CAR T cell Infusion Date [Day 0-R]'] = pd.to_datetime(merged_df['CAR T cell Infusion Date [Day 0-R]']).dt.strftime('%m/%d/%Y')

                #INITLF Retx
                if 'DSINITLF' in self.data:
                    # Subset and rename columns simultaneously
                    DSINITLF_cols_to_select = {
                        'Subject': 'Subject',
                        'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)': 'Phase Entering LTFU',
                        'Last Study Visit Completed in Retreatment (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCRETX_cl_NS_LVCPFUR1)': 'Last Study Visit Completed in Retreatment',
                        'End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)': 'Initiation of Retreatment LTFU Date'
                    }
                    DSINITLF_df = self.data['DSINITLF'][list(DSINITLF_cols_to_select.keys())].rename(columns=DSINITLF_cols_to_select)

                    # Filter rows and drop unnecessary column
                    DSINITLF_df = DSINITLF_df[DSINITLF_df['Phase Entering LTFU'] == 'Retreatment'].drop('Phase Entering LTFU', axis=1)

                    # Merge dataframes
                    merged_df = pd.merge(merged_df, DSINITLF_df, on='Subject', how='left')

                    # Format date
                    merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')

                
                # EOS
                if 'DSEOS' in self.data:
                    DSEOS_df = self.data['DSEOS'][['Subject','End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)' ]].copy()
                    DSEOS_new_col_name = {'End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)':'End of Study Date'}
                    DSEOS_df = DSEOS_df.rename(columns=DSEOS_new_col_name)
                    merged_df = pd.merge(merged_df, DSEOS_df, on='Subject', how='left')
                    merged_df['End of Study Date'] = pd.to_datetime(merged_df['End of Study Date']).dt.strftime('%m/%d/%Y')
            
                # Remove all duplicates
                merged_df = merged_df.drop_duplicates(subset=['Subject'])
                #update headers and fill N/A
                merged_df = merged_df.rename(columns={'Subject': 'Subject ID#'})
                merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna())] = merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna())].fillna('Missing Data')
                merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
                
                return merged_df

            if self.study_name =='12423':
                # DM
                # Check missing 
                if 'DM' in self.data:
                    DM_df = self.data['DM'][['Subject','Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)', 'Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)', 'Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)', 'Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)', 'Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)', 'Apheresis Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)' ]].copy()
                    DM_new_col_name = {'Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)': 'Race', 'Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)': 'Ethnicity', 'Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)': 'Sex Assigned at Birth' , 'Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)': 'Legal Sex', 'Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)': 'Date of Birth', 'Apheresis Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)' : 'Apheresis Consent Date'}
                    DM_df = DM_df.rename(columns=DM_new_col_name)
                    DM_df['Apheresis Consent Date'] = pd.to_datetime(DM_df['Apheresis Consent Date'])
                    sorted_DM_df = DM_df.sort_values(['Subject'])

                # DSCA
                if 'DSCA' in self.data:
                    DSCA_df = self.data['DSCA'][['Subject','Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)' ]].copy()
                    DSCA_new_col_name = {'Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)':'Cohort'}
                    DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
                    merged_df = pd.merge(sorted_DM_df, DSCA_df, on='Subject', how='left')
                    index_reference = merged_df.columns.get_loc('Race')
                    merged_df.insert(index_reference, 'Cohort', merged_df.pop('Cohort'))
                    self.DSCA_exist = 1
                    # print(merged_df)
                else:
                    merged_df = sorted_DM_df
                    self.DSCA_exist = 0
                    

                #DSDLA
                if 'DSDLA' in self.data:
                    DSDLA_df = self.data['DSDLA'][['Subject','Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)' ]].copy()
                    DSDLA_new_col_name = {'Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)':'Assigned Dose Level'}
                    DSDLA_df = DSDLA_df.rename(columns=DSDLA_new_col_name)
                    merged_df = pd.merge(merged_df, DSDLA_df, on='Subject', how='left')
                    index_reference = merged_df.columns.get_loc('Race')
                    merged_df.insert(index_reference, 'Assigned Dose Level', merged_df.pop('Assigned Dose Level'))

                # IE
                if 'IE' in self.data:
                    # Subset and rename columns simultaneously
                    IE_cols_to_select = {
                        'Subject': 'Subject',
                        'Main Consent Date (IG_NS_NA_IE1.DT_NS_NH_MAINCDAT)': 'Main Consent Date',
                        'Subject Meets All Study Eligibility (IG_NS_NA_IE3.CL_NS_YH_ELIGYN_cl_YS_YN1)': 'Subject Meets All Study Eligibility',
                        'Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)': 'Date Physician-Investigator Confirmed Eligibility',
                        'Date of Completion of Monitoring Visit for Eligibility (IG_NS_NA_IE5.DT_NS_YH_ELIGMONDAT)': 'Date of Monitoring Visit for Eligibility'
                    }
                    IE_df = self.data['IE'][list(IE_cols_to_select.keys())].rename(columns=IE_cols_to_select)
                    # Filter Subject Meets All Study Eligibility == Yes
                    IE_df = IE_df[IE_df['Subject Meets All Study Eligibility'] == 'Yes'].drop('Subject Meets All Study Eligibility', axis=1)
                    #convert date columsn to datetime type
                    IE_df['Date Physician-Investigator Confirmed Eligibility'] = pd.to_datetime(IE_df['Date Physician-Investigator Confirmed Eligibility']).dt.strftime('%m/%d/%Y')
                    IE_df['Date of Monitoring Visit for Eligibility'] = pd.to_datetime(IE_df['Date of Monitoring Visit for Eligibility']).dt.strftime('%m/%d/%Y')
                    merged_df = pd.merge(merged_df, IE_df, on='Subject', how='left')
                    merged_df['Main Consent Date'] = pd.to_datetime(merged_df['Main Consent Date'])
                    merged_df['Date of Birth'] = pd.to_datetime(merged_df['Date of Birth'])
                    # Calculate time difference in days and convert to years
                    mask = ~merged_df[['Apheresis Consent Date', 'Date of Birth']].isnull().any(axis=1)
                    # apply relativedelta only to rows with non-NaT values in both columns
                    merged_df.loc[mask, 'Age at Consent'] = merged_df[mask].apply(lambda x: relativedelta(x['Apheresis Consent Date'], x['Date of Birth']).years, axis=1)
                    # for rows that 'Apheresis Consent Date' isnull but 'Main Consent Date' is not null, then use 'Main Consent Date' instead to calculate age
                    merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull()), 'Age at Consent'] = merged_df.loc[(merged_df['Apheresis Consent Date'].isnull() & merged_df['Main Consent Date'].notnull())].apply(lambda x: relativedelta(x['Main Consent Date'], x['Date of Birth']).years, axis=1)
                    merged_df['Apheresis Consent Date'] = merged_df['Apheresis Consent Date'].dt.strftime('%m/%d/%Y')
                    merged_df['Main Consent Date'] = merged_df['Main Consent Date'].dt.strftime('%m/%d/%Y')
                    merged_df = merged_df.drop('Date of Birth', axis = 1)
                    #move the column
                    index_reference = merged_df.columns.get_loc('Apheresis Consent Date')
                    merged_df.insert(index_reference, 'Age at Consent', merged_df.pop('Age at Consent'))

                #APH
                if 'PRAPH' in self.data:
                    PRAPH_cols_to_select = {
                        'Subject': 'Subject',
                        'Event Group Label': 'Event Group Label',
                        'Apheresis Type (IG_NS_NA_PRAPH1.CL_NS_YH_APHTP_cl_NS_APHTP1)': 'Apheresis Type (Fresh or Historical)',
                        'Apheresis Date (IG_NS_NA_PRAPH1.DT_NS_NH_APHDAT)': 'Date of Apheresis Collection'
                    }
                    PRAPH_df = self.data['PRAPH'][list(PRAPH_cols_to_select.keys())].rename(columns=PRAPH_cols_to_select)

                    # Filter rows and drop unnecessary column
                    PRAPH_df = PRAPH_df[PRAPH_df['Event Group Label'] == 'Initial Study Enrollment/Apheresis'].drop('Event Group Label', axis=1)
                    merged_df = pd.merge(merged_df, PRAPH_df, on='Subject', how='left')
                    merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')

                #EXCHMO
                if 'EXCHMO' in self.data:
                    EXCHMO_df = self.data['EXCHMO'][['Subject','Event Group Label', 'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)']].copy()
                    EXCHMO_df = EXCHMO_df[EXCHMO_df['Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)'] != 'NaN']
                    EXCHMO_df = EXCHMO_df[EXCHMO_df['Event Group Label'] == 'Lymphodepleting Chemotherapy']
                    EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject'])
                    EXCHMO_new_col_name = {'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)': 'Date of Initiation of LD Chemo'}
                    EXCHMO_df = EXCHMO_df.rename(columns=EXCHMO_new_col_name)
                    EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                    EXCHMO_df = EXCHMO_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                    merged_df['Date of Initiation of LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of LD Chemo']).dt.strftime('%m/%d/%Y')
                

                #INF
                if 'EXINF' in self.data:
                    EXINF_df = self.data['EXINF'][['Subject','Event Group Label', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)']].copy()
                    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] != 'Day 0-R']
                    EXINF_new_col_name = {'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)': 'CAR T cell Infusion Date [Day 0]'}
                    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                    merged_df['CAR T cell Infusion Date [Day 0]'] = pd.to_datetime(merged_df['CAR T cell Infusion Date [Day 0]']).dt.strftime('%m/%d/%Y')

                #DSINITLF
                if 'DSINITLF' in self.data:
                    DSINITLF_cols_to_select = {
                        'Subject': 'Subject',
                        'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)': 'Phase Entering LTFU',
                        'Last Study Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCPFU_cl_YS_LVCPFU1)': 'Last Study Visit Completed in Primary Follow-Up',
                        'End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)': 'Initiation of LTFU Date'
                    }
                    DSINITLF_df = self.data['DSINITLF'][list(DSINITLF_cols_to_select.keys())].rename(columns=DSINITLF_cols_to_select)

                    # Filter rows and drop unnecessary column
                    DSINITLF_df = DSINITLF_df[DSINITLF_df['Phase Entering LTFU'] == 'Primary Follow-Up'].drop('Phase Entering LTFU', axis=1)
                    # Merge dataframes
                    merged_df = pd.merge(merged_df, DSINITLF_df, on='Subject', how='left')
                    # Format date
                    merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')
                    # print(merged_df)

                #DSINITRT
                if 'DSINITRT' in self.data:
                    DSINITRT_cols_to_select = {
                        'Subject': 'Subject',
                        'Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_YS_LVCPFU1)': 'Last Visit Completed in PFU',
                        'From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)': 'Phase',
                        'End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)': 'Initiation of Retx Date',
                        'End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)': 'End of Long-Term Follow-Up Date'
                    }
                    DSINITRT_df = self.data['DSINITRT'][list(DSINITRT_cols_to_select.keys())].rename(columns=DSINITRT_cols_to_select)
                    merged_df = pd.merge(merged_df, DSINITRT_df, on='Subject', how='left')
                    # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
                    merged_df.loc[merged_df['Phase'] == 'Primary Follow-Up', 'Initiation of LTFU Date'] = 'N/A'
                    merged_df['Last Study Visit Completed in Primary Follow-Up'] = merged_df['Last Study Visit Completed in Primary Follow-Up'].fillna(merged_df['Last Visit Completed in PFU'])
                    merged_df = merged_df.drop('Last Visit Completed in PFU', axis=1)
                    merged_df['Initiation of Retx Date'] = merged_df['Initiation of Retx Date'].fillna(merged_df['End of Long-Term Follow-Up Date'])
                    merged_df = merged_df.drop(['End of Long-Term Follow-Up Date','Phase'], axis=1)
                    # drop End of Long-Term Follow-Up Date column
                    merged_df['Initiation of Retx Date'] = pd.to_datetime(merged_df['Initiation of Retx Date']).dt.strftime('%m/%d/%Y')

                #EXCHMO Retreatment
                if 'EXCHMO' in self.data:
                    # Subset and rename columns simultaneously
                    EXCHMO_cols_to_select = {
                        'Subject': 'Subject',
                        'Event Group Label': 'Event Group Label',
                        'Start Date (IG_NS_NA_EXCHMO2.DT_NS_NH_EXSTDAT)': 'Date of Initiation of Retx LD Chemo'
                    }
                    EXCHMO_df = self.data['EXCHMO'][list(EXCHMO_cols_to_select.keys())].rename(columns=EXCHMO_cols_to_select)

                    # Filter rows
                    EXCHMO_df = EXCHMO_df[
                        (EXCHMO_df['Date of Initiation of Retx LD Chemo'] != 'NaN') & 
                        (EXCHMO_df['Event Group Label'] == 'Retreatment Lymphodepleting Chemotherapy')
                    ]

                    # Drop duplicates and unnecessary columns
                    EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject']).drop('Event Group Label', axis=1)
                    # Sort by Subject
                    EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                    merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                    merged_df['Date of Initiation of Retx LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of Retx LD Chemo']).dt.strftime('%m/%d/%Y')
                

                # EXINF Retx
                if 'EXINF' in self.data:
                    EXINF_df = self.data['EXINF'][['Subject','Event Group Label', 'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)']].copy()
                    EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0-R']
                    EXINF_new_col_name = {'Infusion Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)': 'CAR T cell Infusion Date [Day 0-R]'}
                    EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                    EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                    merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                    merged_df['CAR T cell Infusion Date [Day 0-R]'] = pd.to_datetime(merged_df['CAR T cell Infusion Date [Day 0-R]']).dt.strftime('%m/%d/%Y')

                #INITLF Retx
                if 'DSINITLF' in self.data:
                    # Subset and rename columns simultaneously
                    DSINITLF_cols_to_select = {
                        'Subject': 'Subject',
                        'From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)': 'Phase Entering LTFU',
                        'Last Study Visit Completed in Retreatment (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCRETX_cl_NS_LVCPFUR1)': 'Last Study Visit Completed in Retreatment',
                        'End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)': 'Initiation of Retreatment LTFU Date'
                    }
                    DSINITLF_df = self.data['DSINITLF'][list(DSINITLF_cols_to_select.keys())].rename(columns=DSINITLF_cols_to_select)

                    # Filter rows and drop unnecessary column
                    DSINITLF_df = DSINITLF_df[DSINITLF_df['Phase Entering LTFU'] == 'Retreatment'].drop('Phase Entering LTFU', axis=1)

                    # Merge dataframes
                    merged_df = pd.merge(merged_df, DSINITLF_df, on='Subject', how='left')

                    # Format date
                    merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')

                
                # EOS
                if 'DSEOS' in self.data:
                    DSEOS_df = self.data['DSEOS'][['Subject','End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)' ]].copy()
                    DSEOS_new_col_name = {'End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)':'End of Study Date'}
                    DSEOS_df = DSEOS_df.rename(columns=DSEOS_new_col_name)
                    merged_df = pd.merge(merged_df, DSEOS_df, on='Subject', how='left')
                    merged_df['End of Study Date'] = pd.to_datetime(merged_df['End of Study Date']).dt.strftime('%m/%d/%Y')
            
                #update headers and fill N/A
                merged_df = merged_df.rename(columns={'Subject': 'Subject ID#'})
                merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna())] = merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna())].fillna('Missing Data')
                merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
                
                return merged_df

            if self.study_name =='15420':
                # DM
                # Check missing 
                if 'DM' in self.data:
                    DM_df = self.data['DM'][['Subject','Race (ig_DM1.RACE)', 'Ethnicity (ig_DM1.ETHNIC)', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)', 'Legal Sex (ig_DM1.SEX)', 'Date of Birth (ig_DM1.BRTHDAT)' ]].copy()
                    DM_new_col_name = {'Race (ig_DM1.RACE)': 'Race', 'Ethnicity (ig_DM1.ETHNIC)': 'Ethnicity', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)': 'Sex Assigned at Birth', 'Legal Sex (ig_DM1.SEX)': 'Legal Sex', 'Date of Birth (ig_DM1.BRTHDAT)': 'Date of Birth'}
                    DM_df = DM_df.rename(columns=DM_new_col_name)
                    sorted_DM_df = DM_df.sort_values(['Subject'])
                else:
                    self.error_message.append('Missing DM')

                # DSCA
                if 'DSCA' in self.data:
                    DSCA_df = self.data['DSCA'][['Subject','Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)' ]].copy()
                    DSCA_new_col_name = {'Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)':'Cohort'}
                    DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
                    merged_df = pd.merge(sorted_DM_df, DSCA_df, on='Subject', how='left')
                    index_reference = merged_df.columns.get_loc('Race')
                    merged_df.insert(index_reference, 'Cohort', merged_df.pop('Cohort'))
                    self.DSCA_exist = 1
                    # print(merged_df)
                elif 'PRDIAG' in self.data:
                    PRDIAG_df = self.data['PRDIAG'][['Subject','Disease Type (ig_PRDIAG1.RSCAT)' ]].copy()
                    PRDIAG_new_col_name = {'Disease Type (ig_PRDIAG1.RSCAT)':'Cohort'}
                    PRDIAG_df = PRDIAG_df.rename(columns=PRDIAG_new_col_name)
                    merged_df = pd.merge(sorted_DM_df, PRDIAG_df, on='Subject', how='left')
                    index_reference = merged_df.columns.get_loc('Race')
                    merged_df.insert(index_reference, 'Cohort', merged_df.pop('Cohort'))
                    self.DSCA_exist = 1
                    # print(merged_df)
                else:
                    merged_df = sorted_DM_df
                    self.DSCA_exist = 0
                    

                # DLA

                DLA_df = self.data['DLA'][['Subject','Dose Level Assignment (ig_DLA1.DLADOSELVL)' ]].copy()
                DLA_new_col_name = {'Dose Level Assignment (ig_DLA1.DLADOSELVL)':'Assigned Dose Level'}
                DLA_df = DLA_df.rename(columns=DLA_new_col_name)
                merged_df = pd.merge(merged_df, DLA_df, on='Subject', how='left')
                index_reference = merged_df.columns.get_loc('Race')
                merged_df.insert(index_reference, 'Assigned Dose Level', merged_df.pop('Assigned Dose Level'))


                # IE
                IE_df = self.data['IE'][['Subject', 'Did subject sign the consent form? (ig_IE1.SIGNMAINC)',  'Consent Date (ig_IE1.MAINCDAT)', 'Date of eligibility confirmation by physician-investigator (ig_IE5.ELIGPIDAT)', 'Date of completion of monitoring visit for eligibility (ig_IE5.ELIGMONDAT)']].copy()
                IE_df = IE_df[IE_df['Did subject sign the consent form? (ig_IE1.SIGNMAINC)'] == 'Yes']
                IE_df = IE_df.drop('Did subject sign the consent form? (ig_IE1.SIGNMAINC)', axis = 1)
                IE_new_col_name = {'Consent Date (ig_IE1.MAINCDAT)':'Consent Date', 'Date of eligibility confirmation by physician-investigator (ig_IE5.ELIGPIDAT)': 'Date Physician-Investigator Confirmed Eligibility', 'Date of completion of monitoring visit for eligibility (ig_IE5.ELIGMONDAT)': 'Date of Monitoring Visit for Eligibility'}
                IE_df = IE_df.rename(columns=IE_new_col_name)
                IE_df['Date Physician-Investigator Confirmed Eligibility'] = pd.to_datetime(IE_df['Date Physician-Investigator Confirmed Eligibility']).dt.strftime('%m/%d/%Y')
                IE_df['Date of Monitoring Visit for Eligibility'] = pd.to_datetime(IE_df['Date of Monitoring Visit for Eligibility']).dt.strftime('%m/%d/%Y')
                merged_df = pd.merge(merged_df, IE_df, on='Subject', how='left')
                #convert date columsn to datetime type
                merged_df['Consent Date'] = pd.to_datetime(merged_df['Consent Date'])
                merged_df['Date of Birth'] = pd.to_datetime(merged_df['Date of Birth'])
                # create a mask for non-NaT values in the two columns
                mask = ~merged_df[['Consent Date', 'Date of Birth']].isnull().any(axis=1)

                # apply relativedelta only to rows with non-NaT values in both columns
                merged_df.loc[mask, 'Age at Consent'] = merged_df[mask].apply(lambda x: relativedelta(x['Consent Date'], x['Date of Birth']).years, axis=1)
                
                merged_df['Consent Date'] = pd.to_datetime(merged_df['Consent Date']).dt.strftime('%m/%d/%Y')
                # print(merged_df)
                merged_df = merged_df.drop('Date of Birth', axis = 1)

                #move the column
                index_reference = merged_df.columns.get_loc('Consent Date')
                merged_df.insert(index_reference, 'Age at Consent', merged_df.pop('Age at Consent'))

                #APH
                APH_df = self.data['APH'][['Subject','Event Group Label','Apheresis Type (ig_APH1.APHCAT)', 'Apheresis Date (ig_APH1.APHDAT)']].copy()
                APH_new_col_name = {'Apheresis Type (ig_APH1.APHCAT)': 'Apheresis Type (Fresh or Historical)', 'Apheresis Date (ig_APH1.APHDAT)': 'Date of Apheresis Collection'}
                APH_df = APH_df[APH_df['Event Group Label'] == 'Apheresis']
                APH_df = APH_df.drop('Event Group Label', axis = 1)
                APH_df = APH_df.rename(columns=APH_new_col_name)
                merged_df = pd.merge(merged_df, APH_df, on='Subject', how='left')
                merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')

                #EXCHMO
                EXCHMO_df = self.data['EXCHMO'][['Subject','Event Group Label', 'Start Date (ig_EXCHMO2.EXSTDAT)']].copy()
                EXCHMO_df = EXCHMO_df[EXCHMO_df['Start Date (ig_EXCHMO2.EXSTDAT)'] != 'NaN']
                EXCHMO_df = EXCHMO_df[EXCHMO_df['Event Group Label'] != 'Retreatment Lymphodepleting Chemotherapy']
                EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject'])
                EXCHMO_new_col_name = {'Start Date (ig_EXCHMO2.EXSTDAT)': 'Date of Initiation of LD Chemo'}
                EXCHMO_df = EXCHMO_df.rename(columns=EXCHMO_new_col_name)
                EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                EXCHMO_df = EXCHMO_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                merged_df['Date of Initiation of LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of LD Chemo']).dt.strftime('%m/%d/%Y')
                

                #INF
                INF_df = self.data['INF'][['Subject','Event Group Label', 'Infusion Date (ig_INF1.INFDAT)']].copy()
                INF_df = INF_df[INF_df['Event Group Label'] != 'Day 0-R']
                INF_new_col_name = {'Infusion Date (ig_INF1.INFDAT)': 'CAR T cell Infusion Date [Day 0]'}
                INF_df = INF_df.rename(columns=INF_new_col_name)
                INF_df = INF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, INF_df, on='Subject', how='left')
                merged_df['CAR T cell Infusion Date [Day 0]'] = pd.to_datetime(merged_df['CAR T cell Infusion Date [Day 0]']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #INITLF
                INITLF_df = self.data['INITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)', 'Last Study Visit Completed in Primary Follow-Up (ig_INITLF1.DSLVCPFU)', 'End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)'] != 'Retreatment']
                INITLF_new_col_name = {'Last Study Visit Completed in Primary Follow-Up (ig_INITLF1.DSLVCPFU)': 'Last Study Visit Completed in Primary Follow-Up', 'End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)': 'Initiation of LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #DSINITRT
                DSINITRT_df = self.data['DSINITRT'][['Subject','Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)', 'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)', 'End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)']].copy()
                DSINITRT_new_col_name = {'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)': 'Retreatment?', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)':'Phase','End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)': 'Initiation of Retx Date'}
                DSINITRT_df = DSINITRT_df.rename(columns=DSINITRT_new_col_name)
                DSINITRT_df['Initiation of Retx Date'] = DSINITRT_df['Initiation of Retx Date'].fillna(DSINITRT_df['End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)'])
                DSINITRT_df = DSINITRT_df.drop('End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)', axis = 1)
                merged_df = pd.merge(merged_df, DSINITRT_df, on='Subject', how='left')
                # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
                merged_df.loc[merged_df['Phase'] == 'Primary Follow-Up', 'Initiation of LTFU Date'] = 'N/A'
                merged_df['Last Study Visit Completed in Primary Follow-Up'] = merged_df['Last Study Visit Completed in Primary Follow-Up'].fillna(merged_df['Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)'])
                merged_df = merged_df.drop(['Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)','Phase'], axis = 1)
                merged_df['Initiation of Retx Date'] = pd.to_datetime(merged_df['Initiation of Retx Date']).dt.strftime('%m/%d/%Y')

                #EXCHMO Retreatment
                EXCHMO_df = self.data['EXCHMO'][['Subject','Event Group Label', 'Start Date (ig_EXCHMO2.EXSTDAT)']].copy()
                EXCHMO_df = EXCHMO_df[EXCHMO_df['Start Date (ig_EXCHMO2.EXSTDAT)'] != 'NaN']
                EXCHMO_df = EXCHMO_df[EXCHMO_df['Event Group Label'] != 'Lymphodepleting Chemotherapy']
                EXCHMO_df = EXCHMO_df.drop_duplicates(subset=['Subject'])
                EXCHMO_new_col_name = {'Start Date (ig_EXCHMO2.EXSTDAT)': 'Date of Initiation of Retx LD Chemo'}
                EXCHMO_df = EXCHMO_df.rename(columns=EXCHMO_new_col_name)
                EXCHMO_df = EXCHMO_df.sort_values(['Subject'])
                EXCHMO_df = EXCHMO_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, EXCHMO_df, on='Subject', how='left')
                merged_df['Date of Initiation of Retx LD Chemo'] = pd.to_datetime(merged_df['Date of Initiation of Retx LD Chemo']).dt.strftime('%m/%d/%Y')
                

                # INF Retx
                INF_df = self.data['INF'][['Subject','Event Group Label', 'Infusion Date (ig_INF1.INFDAT)']].copy()
                INF_df = INF_df[INF_df['Event Group Label'] == 'Day 0-R']
                INF_new_col_name = {'Infusion Date (ig_INF1.INFDAT)': 'CAR T cell Retreatment Date [Day 0-R]'}
                INF_df = INF_df.rename(columns=INF_new_col_name)
                INF_df = INF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, INF_df, on='Subject', how='left')
                merged_df['CAR T cell Retreatment Date [Day 0-R]'] = pd.to_datetime(merged_df['CAR T cell Retreatment Date [Day 0-R]']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #INITLF Retx
                INITLF_df = self.data['INITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)', 'Last Study Visit Completed in Retreatment (ig_INITLF1.DSLVCRETX)', 'End of Retreatment Date (ig_INITLF1.DSENRETXDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)'] == 'Retreatment']
                INITLF_new_col_name = {'Last Study Visit Completed in Retreatment (ig_INITLF1.DSLVCRETX)': 'Last Study Visit Completed in Retreatment F/up', 'End of Retreatment Date (ig_INITLF1.DSENRETXDAT)': 'Initiation of Retreatment LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')
                
                # EOS
                EOS_df = self.data['EOS'][['Subject','End of Study Date (ig_EOS1.EOSDAT)' ]].copy()
                EOS_new_col_name = {'End of Study Date (ig_EOS1.EOSDAT)':'End of Study Date'}
                EOS_df = EOS_df.rename(columns=EOS_new_col_name)
                merged_df = pd.merge(merged_df, EOS_df, on='Subject', how='left')
                merged_df['End of Study Date'] = pd.to_datetime(merged_df['End of Study Date']).dt.strftime('%m/%d/%Y')
            
                #update headers and fill N/A
                merged_df = merged_df.rename(columns={'Subject': 'Subject ID#'})
                merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna()), ['Retreatment?']] = merged_df.loc[(merged_df['End of Study Date'].notna() & merged_df['CAR T cell Infusion Date [Day 0]'].notna()), ['Retreatment?']].fillna('Missing Data')
                merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
                
                return merged_df
            
            if self.study_name =='16321':
                # DM
                # if 'DM' in self.data:
                DM_df = self.data['DM'][['Subject','Race (ig_DM1.RACE)', 'Ethnicity (ig_DM1.ETHNIC)', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)', 'Legal Sex (ig_DM1.SEX)', 'Date of Birth (ig_DM1.BRTHDAT)', 'Apheresis Consent Date (ig_DM1.RFICDAT)' ]].copy()
                DM_new_col_name = {'Race (ig_DM1.RACE)': 'Race', 'Ethnicity (ig_DM1.ETHNIC)': 'Ethnicity', 'Sex Assigned at Birth (ig_DM1.BRTHSEX)': 'Sex Assigned at Birth', 'Legal Sex (ig_DM1.SEX)': 'Legal Sex', 'Date of Birth (ig_DM1.BRTHDAT)': 'Date of Birth', 'Apheresis Consent Date (ig_DM1.RFICDAT)' : 'Apheresis Consent Date'}
                DM_df = DM_df.rename(columns=DM_new_col_name)
                sorted_DM_df = DM_df.sort_values(['Subject'])

                # DSCA

                DSCA_df = self.data['DSCA'][['Subject','Cohort Assignment (ig_DSCA1.CACHASCOD)' ]].copy()
                DSCA_new_col_name = {'Cohort Assignment (ig_DSCA1.CACHASCOD)':'Assigned Cohort'}
                DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
                merged_df = pd.merge(sorted_DM_df, DSCA_df, on='Subject', how='left')
                index_reference = merged_df.columns.get_loc('Race')
                merged_df.insert(index_reference, 'Assigned Cohort', merged_df.pop('Assigned Cohort'))
                # print(merged_df)


                # IE
                IE_df = self.data['IE'][['Subject',  'Main Consent Date (ig_IE1.MAINCDAT)', 'Date of Eligibility Confirmation by Physician-Investigator (ig_IE5.ELIGPIDAT)', 'Date of Completion of Monitoring Visit for Eligibility (ig_IE5.ELIGMONDAT)']].copy()
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
                APH_df = self.data['PRAPH'][['Subject','Apheresis Type (ig_PRAPH1.APHCAT)', 'Apheresis Date (ig_PRAPH1.APHDAT)']].copy()
                APH_new_col_name = {'Apheresis Type (ig_PRAPH1.APHCAT)': 'Apheresis Type (Fresh or Historical)', 'Apheresis Date (ig_PRAPH1.APHDAT)': 'Date of Apheresis Collection'}
                APH_df = APH_df.rename(columns=APH_new_col_name)
                merged_df = pd.merge(merged_df, APH_df, on='Subject', how='left')
                merged_df['Date of Apheresis Collection'] = pd.to_datetime(merged_df['Date of Apheresis Collection']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #EXINF
                EXINF_df = self.data['EXINF'][['Subject', 'Event Group Label', 'Study Treatment Date (ig_EXINF1.INFDAT)']].copy()
                EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0']
                EXINF_new_col_name = {'Study Treatment Date (ig_EXINF1.INFDAT)': 'CART T cell Administration Date (Day 0)'}
                EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                merged_df['CART T cell Administration Date (Day 0)'] = pd.to_datetime(merged_df['CART T cell Administration Date (Day 0)']).dt.strftime('%m/%d/%Y')
                # print(merged_df)
                

                #DSINITLF
                INITLF_df = self.data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] != 'Retreatment']
                INITLF_new_col_name = {'Last Study Visit Completed in Primary Follow-Up (ig_DSINITLF1.DSLVCPFU)': 'Last Study Visit Completed in Primary Follow-Up', 'End of Primary Follow-Up Date (ig_DSINITLF1.DSENPFUDAT)': 'Initiation of LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of LTFU Date'] = pd.to_datetime(merged_df['Initiation of LTFU Date']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #DSINITRT
                DSINITRT_df = self.data['DSINITRT'][['Subject','Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)', 'Will the Subject receive Retreatment? (ig_DSINITRT1.DSRTYN)', 'From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)', 'End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)', 'End of Long-Term Follow-Up Date (ig_DSINITRT1.DSRENLTFUDAT)']].copy()
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
                EXINF_df = self.data['EXINF'][['Subject','Event Group Label', 'Study Treatment Date (ig_EXINF1.INFDAT)']].copy()
                EXINF_df = EXINF_df[EXINF_df['Event Group Label'] == 'Day 0-R1']
                EXINF_new_col_name = {'Study Treatment Date (ig_EXINF1.INFDAT)': 'CAR T Cell Retreatment Date (Day 0-R1)'}
                EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
                EXINF_df = EXINF_df.drop('Event Group Label', axis = 1)
                merged_df = pd.merge(merged_df, EXINF_df, on='Subject', how='left')
                merged_df['CAR T Cell Retreatment Date (Day 0-R1)'] = pd.to_datetime(merged_df['CAR T Cell Retreatment Date (Day 0-R1)']).dt.strftime('%m/%d/%Y')
                # print(merged_df)

                #INITLF Retx
                INITLF_df = self.data['DSINITLF'][['Subject','From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', 'Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)', 'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)']].copy()
                INITLF_df = INITLF_df[INITLF_df['From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)'] == 'Retreatment']
                INITLF_df = INITLF_df[INITLF_df['Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)'] == 'Retreatment-R1']
                INITLF_new_col_name = {'Last Study Visit Completed in Retreatment (ig_DSINITLF1.DSLVCRETX)': 'Last Study Visit Completed in Retreatment (-R1)', 'End of Retreatment Date (ig_DSINITLF1.DSENRETXDAT)': 'Initiation of Retreatment LTFU Date'}
                INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
                INITLF_df = INITLF_df.drop('From which Phase is the Subject entering Long-Term Follow-Up? (ig_DSINITLF1.DSPHASE)', axis = 1)
                INITLF_df = INITLF_df.drop('Retreatment Cycle Number (ig_DSINITLF1.RETXCYCLEINITLT)', axis = 1)
                merged_df = pd.merge(merged_df, INITLF_df, on='Subject', how='left')
                merged_df['Initiation of Retreatment LTFU Date'] = pd.to_datetime(merged_df['Initiation of Retreatment LTFU Date']).dt.strftime('%m/%d/%Y')
                
                # EOS
                EOS_df = self.data['DSEOS'][['Subject','End of Study Date (ig_DSEOS1.EOSDAT)' ]].copy()
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
            

    def output(self):

            with pd.ExcelWriter(self.output_dir  + '/' + self.output_file_name + '.xlsx') as writer:  
                self.output_df.to_excel(writer, sheet_name='Enrollment Log ' + self.study_name, index = False)
                worksheet = writer.sheets['Enrollment Log ' + self.study_name]
                # worksheet.set_column(0, self.output_df.shape[1]-1, 15)
                worksheet.autofit()
                border_format = writer.book.add_format({'border': 2, 'text_wrap': True, 'align': 'left'})
                blue_header_format = writer.book.add_format({'bg_color': '#B7DEE8',
                                                                'text_wrap': True,
                                                                'valign': 'vcenter',
                                                                'align': 'center',
                                                                'bold' : True,
                                                                'border': 2})
                purple_header_format = writer.book.add_format({'bg_color': '#7030A0',
                                                                'text_wrap': True,
                                                                'valign': 'vcenter',
                                                                'align': 'center',
                                                                'bold' : True,
                                                                'font_color': 'white',
                                                                'border': 2})
                green_header_format = writer.book.add_format({'bg_color': '#C4D79B',
                                                                'text_wrap': True,
                                                                'valign': 'vcenter',
                                                                'align': 'center'
                                                                ,'bold' : True,
                                                                'border': 2})
                pink_header_format = writer.book.add_format({'bg_color': '#FABF8F',
                                                                'text_wrap': True,
                                                                'valign': 'vcenter',
                                                                'align': 'center',
                                                                'bold' : True,
                                                                'border': 2})
                yellow_header_format = writer.book.add_format({'bg_color': '#FFFF00',
                                                                'text_wrap': True,
                                                                'valign': 'vcenter',
                                                                'align': 'center',
                                                                'bold' : True,
                                                                'border': 2})
                border_format = writer.book.add_format({'border': 1})
                worksheet.conditional_format(0, 0, len(self.output_df.index), len(self.output_df.columns)-1, {'type': 'formula', 'criteria': 'True', 'format': border_format})
                if self.study_name == '11823':
                    for i in range (7):
                        worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                    for i in range (7, 11):
                        worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                    for i in range (11, 16):
                        worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                    for i in range (16, 22 ):
                        worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                    for i in range (22, 23):
                        worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                if self.study_name == '12423':
                    for i in range (7 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                    for i in range (7 + self.DSCA_exist, 11 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                    for i in range (11 + self.DSCA_exist, 16 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                    for i in range (16 + self.DSCA_exist, 22 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                    for i in range (22 + self.DSCA_exist, 23 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                if self.study_name == '15420':
                    for i in range (7 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                    for i in range (7 + self.DSCA_exist, 10 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                    for i in range (10 + self.DSCA_exist, 16 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                    for i in range (16 + self.DSCA_exist, 22 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                    for i in range (22 + self.DSCA_exist, 23 + self.DSCA_exist):
                        worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                elif self.study_name == '16321':
                    for i in range (7):
                        worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                    for i in range (7,11):
                        worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                    for i in range (11,16):
                        worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                    for i in range (16,21):
                        worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                    for i in range (21,22):
                        worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                elif self.study_name == '03821':
                    for i in range (8):
                        worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                    for i in range (8,12):
                        worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                    for i in range (12,18):
                        worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                    for i in range (18,23):
                        worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                    for i in range (23,24):
                        worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                
                worksheet.set_row(0, 60)
                worksheet.set_column('A:AA', 15)
            # writer.save()

if __name__ == '__main__':
    EnrollmentLog("15420", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_15420_huCART19-IL18_2024_01_03_08_23_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Download", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
