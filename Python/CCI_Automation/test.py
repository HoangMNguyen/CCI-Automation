# This Python file uses the following encoding: utf-8
from util import *
from pathlib import Path
import sys
from datetime import date
from EnrollmentLog.EnrollmentLog import EnrollmentLog
from Clockify.Clockify import ClockifyDashboard
from DSMB.DSMB import DSMB
from DSMB.DSMB_util import *


if __name__ == "__main__":

    # Run DSMB
    DSMB("15420", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_15420_huCART19-IL18_PPT2_2024_02_16_15_27_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", "test original 2-16-2024")
    # EnrollmentLog("03821", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_03821_M5-VCN_DEV1_2024_01_25_13_15_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # EnrollmentLog("11823", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_11823_TmPSMA-02_DEV1_2024_01_25_10_08_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # EnrollmentLog("12423", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_12423-TmCD19-IL18_2024_01_24_08_45_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # EnrollmentLog("15420", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_15420_huCART19-IL18_PPT2_2024_01_11_08_58_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # EnrollmentLog("16321", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_16321_EGFR-IL13Ra2_PPT2_2024_01_25_13_31_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # DSMB("12423", "C:/Users/hmn39/Dropbox/Current Work/Download/Core_Listings_12423-TmCD19-IL18_2024_02_13_10_08_EST.zip", "C:/Users/hmn39/Dropbox/Current Work/Output", "test1", debug=True)


    # Run DSMB 
    # DSMB("15420", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Download/Core_Listings_15420_huCART19-IL18_PPT2_2024_02_16_15_27_EST.zip", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Output", "test original 2-16-2024")
    # EnrollmentLog("03821", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Download/Core_Listings_03821_M5-VCN_DEV1_2024_01_25_13_15_EST.zip", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Output", datetime.now().strftime("%Y%m%d%H%M%S") + "test" )
    # DSMB("12423", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Download/Core_Listings_12423-TmCD19-IL18_PPT1_2024_02_12_09_04_EST.zip", "C:/Users/Hoang Nguyen/Dropbox/Current Work/Output", debug=True)

