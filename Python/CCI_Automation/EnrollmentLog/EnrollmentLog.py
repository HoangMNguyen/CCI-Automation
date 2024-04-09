#!/usr/bin/env python3
import pandas as pd
from dateutil.relativedelta import *
from util import *


class EnrollmentLog():
    def __init__(self, study_name, input_dir, output_dir, output_file_name, cut_off_date = None):
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
        if self.study_name == '12423':
            from .EnrollmentLog12423 import EnrollmentLog12423
            output_df = EnrollmentLog12423(self.data)
        elif self.study_name == '03821':
            from .EnrollmentLog03821 import EnrollmentLog03821
            output_df = EnrollmentLog03821(self.data)  
        elif self.study_name =='11823':
            from .EnrollmentLog11823 import EnrollmentLog11823
            output_df = EnrollmentLog11823(self.data)
        elif self.study_name =='15122':
            from .EnrollmentLog15122 import EnrollmentLog15122
            output_df = EnrollmentLog15122(self.data)    
        elif self.study_name =='15420':
            from .EnrollmentLog15420 import EnrollmentLog15420
            output_df = EnrollmentLog15420(self.data) 
        elif self.study_name =='16321':
            from .EnrollmentLog16321 import EnrollmentLog16321
            output_df = EnrollmentLog16321(self.data)
        return output_df

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
                    elif self.study_name == '12423':
                        for i in range (8):
                            worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                        for i in range (8, 12):
                            worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                        for i in range (12, 17):
                            worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                        for i in range (17, 23):
                            worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                        for i in range (23, 24):
                            worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                    elif self.study_name == '15122':
                        for i in range (8):
                            worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                        for i in range (8, 12):
                            worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                        for i in range (12, 17):
                            worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                        for i in range (17, 18):
                            worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                        for i in range (18, 19):
                            worksheet.write(0, i, self.output_df.columns.values[i], yellow_header_format)
                    elif self.study_name == '15420':
                        for i in range (8):
                            worksheet.write(0, i, self.output_df.columns.values[i], blue_header_format)
                        for i in range (8, 11):
                            worksheet.write(0, i, self.output_df.columns.values[i], purple_header_format)
                        for i in range (11, 17):
                            worksheet.write(0, i, self.output_df.columns.values[i], green_header_format)
                        for i in range (17, 23):
                            worksheet.write(0, i, self.output_df.columns.values[i], pink_header_format)
                        for i in range (23, 24):
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