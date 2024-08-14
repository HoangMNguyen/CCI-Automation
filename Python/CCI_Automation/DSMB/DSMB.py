#!/usr/bin/env python3
import pandas as pd
import numpy as np
from util import *
from DSMB.DSMB_util import *
from dateutil.relativedelta import *
from datetime import datetime, date
from typing import Optional


class DSMB:
    def __init__(
        self,
        study_name: str,
        input_dir: str,
        output_dir: str,
        output_file_name: str,
        cutoff_date: Optional[date] = None,
        debug: Optional[bool] = False,
    ):
        """DSMB for any study

        Args:
            study_name (string): study selected for the DSMB
            input_dir (string): core listing folder folder/directory path
            output_dir (string): file path where the report is gonna be saved
            Optional cutoff_date (date): cutoff date for the report
        """
        if input_dir == None:
            print("No dir selected!")
        else:
            self.debug = debug
            self.study_name = study_name
            self.output_dir = output_dir
            self.output_file_name = output_file_name
            self.input_data = read_data_dict_zip_corelisting(input_dir, cutoff_date)

            self.enrollment = True
            self.infusion = True
            self.response = True
            self.data_analysis(self.input_data, self.study_name)

    def data_analysis(self, data, study_name):
        pd.set_option("display.max_columns", None)
        pd.set_option("display.max_rows", None)
        if study_name == "12423":
            from .DSMB12423 import DSMB12423

            DSMB12423(
                data,
                export=True,
                output_dir=self.output_dir,
                output_file_name=self.output_file_name,
                debug=self.debug,
            )

        if study_name == "15122":
            from .DSMB15122 import DSMB15122

            DSMB15122(
                data,
                export=True,
                output_dir=self.output_dir,
                output_file_name=self.output_file_name,
                debug=self.debug,
            )
        if study_name == "03821":
            from .DSMB03821 import DSMB03821

            DSMB03821(
                data,
                export=True,
                output_dir=self.output_dir,
                output_file_name=self.output_file_name,
                debug=self.debug,
            )

        if study_name == "15420":
            from .DSMB15420 import DSMB15420

            DSMB15420(
                data,
                export=True,
                output_dir=self.output_dir,
                output_file_name=self.output_file_name,
                debug=self.debug,
            )
