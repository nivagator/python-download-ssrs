# https://www.reddit.com/r/Python/comments/67xak4/enterprise_intranet_authentication_and_ssrs/?st=JDZ961EE&sh=7bff6592

import requests, os, datetime, configparser, pyodbc, logging
from requests import Session
from requests_ntlm import HttpNtlmAuth

def main():
    config = configparser.ConfigParser()
    config.read('sineserver.ini')
    username = config['sineserver']['username']
    password = config['sineserver']['password']
    server = config['sineserver']['server']
    db = config['sineserver']['database']

    py_file_path = os.path.dirname(os.path.realpath(__file__))
    fn = py_file_path + '\\logs\\ssrs_downloads_' + datetime.datetime.now().strftime("%Y_%m_%d") + '.log'

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    # create file handler
    if logger.handlers:
        logger.handlers = []
    # handler = logging.FileHandler('ssrs_download.log')
    handler = logging.FileHandler(fn)
    handler.setLevel(logging.INFO)

    # create logging format
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%m-%d-%Y %I:%M:%S %p')
    handler.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(handler)
    logger.info("-"*40)
    logger.info('Started')   

    rpt_pass = []
    rpt_fail = []
    try:
        cnxn = pyodbc.connect('Driver={ODBC Driver 13 for SQL Server};'
                            r'Server='+ server +';'
                            r'Database='+ db +';'
                            r'UID='+ username +';'
                            r'PWD='+ password +';')
        logger.info("SQL connection complete.")
    except:
        logger.info("SQL connection failed")
        return False

    # PULL DATA FROM SQL SERVER
    cursor = cnxn.cursor()
    cursor.execute("""
                    SELECT
                        RPT_EXPRT_DIM_ID
                        ,[RPT_NM]
                        ,[RPT_FILE_NM] = RPT_FILE_NM 
                            + CAST(year(getdate()) as varchar(4))
                            +'-'+case when len(month(getdate())) = 1 THEN '0' + cast(month(getdate()) as varchar(2)) ELSE CAST(month(getdate()) as varchar(2)) END
                            +'-' +case when len(day(getdate())) = 1 THEN '0' + cast(day(getdate()) as varchar(2))  ELSE CAST(day(getdate()) as varchar(2)) END 
                        ,[RPT_CONN_NM]
                        ,[RPT_SRVR_LOC_DESC]
                        ,[RPT_EXPRT_SUFFIX_DESC]
                        ,[RPT_EXPRT_FILE_TYP] = CASE WHEN [RPT_EXPRT_FILE_TYP] = '.xls' THEN 'excel' ELSE 'pdf' END
                    FROM [ICG_WORK].[work].[RPT_EXPRT_DIM]
                    WHERE 1=1
                    AND ACTV_IND = 1
                    AND EXPIR_DT IS NULL
                    """)
    logger.info("SQL cursor executed")

    # EXPORT REPORTS
    for row in cursor:
        try:
            get_rpt(row.RPT_EXPRT_FILE_TYP,row.RPT_SRVR_LOC_DESC,row.RPT_FILE_NM)
            rpt_pass.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM])
            # print("Download complete: %r" % row.RPT_NM)
            # logger.info("Download complete: %r" % row.RPT_NM)
        except:
            rpt_fail.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM])
            logger.warning("Download failed: %r" % row.RPT_NM)
            # return False
    
    # LOG OUTPUT
    pass_rpts = len(rpt_pass)
    fail_rpts = len(rpt_fail)
    total_rpts = pass_rpts + fail_rpts

    # pass reports
    logger.info("{0} of {1} reports downloaded successfully.".format(pass_rpts,total_rpts))
    if pass_rpts > 0:
        for rpt in rpt_pass:
            logger.info("\tSUCCESS! ID: {0} \t| Name: {1}".format(rpt[0],rpt[1]))

    # fail reports
    logger.info("{0} of {1} reports failed to download.".format(fail_rpts,total_rpts))
    if fail_rpts > 0:
        for rpt in rpt_fail:
            logger.info("\tFAIL! ID: {0} \t| Name: {1}".format(rpt[0],rpt[1]))
    
    logger.info('Finished')
    

def get_rpt(file_format, report_loc, file_name):
    config = configparser.ConfigParser()
    config.read('ssrs_config.ini')
    username = config['DEFAULT']['username']
    password = config['DEFAULT']['password']
    base_url = config['DEFAULT']['base_url']
    # print('u: ' + username, ' | pw: ' + password)

    py_file_path = os.path.dirname(os.path.realpath(__file__))
    dt_str = datetime.datetime.now().strftime("%Y_%m_%d_%H%M%S")

    default_report_loc = r'%2fRevenue%20Reporting%2fBranch+Revenue+Report&BRANCH=222-Team Nunzio'
    default_file_name = 'report_export_' + dt_str
    
    if file_format == 'pdf':
        file_ext = '.pdf'
        rpt_render = '&rs:Format=PDf'
    elif file_format == 'excel':
        file_ext = '.xls'
        rpt_render = '&rs:Format=EXCEL'
    else:
        return('invalid file format')

    url = base_url + default_report_loc + rpt_render
    if report_loc:
        url = base_url + report_loc +rpt_render
    # print(url)

    out_file = py_file_path + '\\' + default_file_name + file_ext
    if file_name:
        out_file = py_file_path + '\\report_exports\\' + file_name + file_ext
    # print(out_file)

    session = Session()
    session.auth = HttpNtlmAuth(username=username, password=password)
    response = session.get(url)
    # print('starting write out')
    try:
        with open(out_file, 'wb') as f:
            f.write(response.content)
    except:
        print('failed to write to file')

    session.close()
    # print('finished')

if __name__ == "__main__":
    main()