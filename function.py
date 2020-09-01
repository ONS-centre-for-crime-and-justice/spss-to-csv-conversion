"""This code runs a python script through a 3.4 interpreter to use IBM's SPSS API"""

from subprocess import Popen

def read_SPSS_into_csv(input_file,
                       output_file,
                       output_var_vals_file,
                       output_var_info_file):
    """
    This function is used to convert SPSS sav files to csv and remove invalid ASCII characters so they can be read into
    python. It also outputs two xls files containing metadata from SPSS.
    :param input_file: (Str) The path to the sav file
    :param output_file: (Str) The path where the csv is to be saved
    :param output_var_vals_file: (Str) The path where the values (Possible values for coded variables in SPSS) is to be
    saved
    :param output_var_info_file: (Str) The path where the variable information (all other metadata) is to be saved
    :return: None
    """
    replace_d = {'user_input_1': input_file,
                 'user_input_2': output_file,
                 'user_input_3': output_var_vals_file,
                 'user_input_4': output_var_info_file}

    # the original file is not altered
    filename_original = r"G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\run_spss_syntax_original.py"
    script = open(filename_original).read()
    # replace the placeholder user_inputs with the paths specified by the arguments
    for key, word in replace_d.items():
        script = script.replace(key, word)

    # this is a blank py file that we write the script to before running it
    filename_production = r"G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\run_spss_syntax_production.py"
    f = open(filename_production, 'w')
    f.write(script)
    f.close()

    Popen([r"C:\Program Files\IBM\SPSS\Statistics\24\Python3\python.exe",filename_production]).wait()

    # delete contents of file for the next run
    open(filename_production, 'w').close()

###############################################################################
################################ VF ########################################## 
###############################################################################


#==============================================================================
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2009-10 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2009-10_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2009-10_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2009-10_info.xls')
#    
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2010-11 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2010-11_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2010-11_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2010-11_info.xls')
#    
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2011-12 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2011-12_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2011-12_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2011-12_info.xls')
#  
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2012-13 - fixes applied.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2012-13_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2012-13_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2012-13_info.xls')
#    
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2013-14 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2013-14_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2013-14_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2013-14_info.xls')
#  
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2014-15 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2014-15_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2014-15_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2014-15_info.xls')
#  
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2015-16 - fixes applied.sav',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2015-16_data.csv',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2015-16_test_values.xls',
#                     r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2015-16_info.xls')
# 
# 
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2016-17 - fixes applied.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2016-17_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2016-17_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2016-17_info.xls')
# 
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2017-18 - fixes applied.sav',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2017-18_data.csv',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2017-18_test_values.xls',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2017-18_info.xls')
# 
#    
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2018-19 - fixes applied.sav',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2018-19_data.csv',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2018-19_test_values.xls',         
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2018-19_info.xls')
#   
# read_SPSS_into_csv(r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\VF_2019-20 - fixes applied.sav',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\VF\VF_2019-20_data.csv',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\VF_2019-20_test_values.xls',
#                      r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\VF_2019-20_info.xls')
#==============================================================================
   

###############################################################################
################################ NVF ########################################## 
###############################################################################
#
#==============================================================================
# read_SPSS_into_csv(r'G:\Final datasets\2009-10\NVF_2009-10.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2009-10_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2009-10_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2009-10_info.xls')
#   
# read_SPSS_into_csv(r'G:\Final datasets\2010-11\NVF_2010-11.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2010-11_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2010-11_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2010-11_info.xls')
#   
# read_SPSS_into_csv(r'G:\Final datasets\2011-12\NVF_2011-12.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2011-12_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2011-12_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2011-12_info.xls')
# 
#   
# read_SPSS_into_csv(r'G:\Final datasets\2012-13\NVF_2012-13.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2012-13_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2012-13_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2012-13_info.xls')
#   
read_SPSS_into_csv(r'G:\Final datasets\2013-14\NVF_2013-14.sav',
                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2013-14_data.csv',
                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2013-14_test_values.xls',
                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2013-14_info.xls')
#  
#  
# read_SPSS_into_csv(r'G:\Final datasets\2014-15\NVF_2014-15.sav',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2014-15_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2014-15_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2014-15_info.xls')
#  
# read_SPSS_into_csv(r'G:\Final datasets\2015-16\NVF_2015-16.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2015-16_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2015-16_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2015-16_info.xls')
#   
# read_SPSS_into_csv(r'G:\Final datasets\2016-17\NVF_2016-17.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2016-17_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2016-17_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2016-17_info.xls')
#
#read_SPSS_into_csv(r'G:\Final datasets\2017-18\NVF_2017-18.sav',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2017-18_data.csv',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2017-18_test_values.xls',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2017-18_info.xls')
# 
#  
# read_SPSS_into_csv(r'G:\Final datasets\2018-19\NVF_2018-19.sav',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2018-19_data.csv',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2018-19_test_values.xls',
#                    r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2018-19_info.xls')
#  
# 
#read_SPSS_into_csv(r'G:\Final datasets\2019-20\NVF_2019-20.sav',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Data\NVF\NVF_2019-20_data.csv',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Values\NVF\NVF_2019-20_test_values.xls',
#                   r'G:\Publications\Nature of Crime Tables 2019-20\NoC Review\Final Datasets\csvs\Info\NVF\NVF_2019-20_info.xls')
  
#==============================================================================

