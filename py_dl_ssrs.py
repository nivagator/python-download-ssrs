import requests, os, datetime, configparser, pyodbc, logging, platform
from requests import Session
from requests_ntlm import HttpNtlmAuth
from send_email import email

def main():
    config = configparser.ConfigParser()
    config.read('config.ini')
    username = config['sineserver']['username']
    password = config['sineserver']['password']
    server = config['sineserver']['server']
    db = config['sineserver']['database']

    sys_type = platform.system()
    py_file_path = os.path.dirname(os.path.realpath(__file__))
    if sys_type == 'Windows':
        filepath = py_file_path + '\\report_exports\\'
        logpath = py_file_path + '\\logs\\'
    else: 
        filepath = py_file_path + '//report_exports//'
        logpath = py_file_path + '//logs//'
    dt = datetime.datetime.now().strftime("%Y_%m_%d")
    logfile = 'ssrs_downloads_' + dt + '.log'
    
    fn =  logpath + logfile

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    # create file handler
    if logger.handlers:
        logger.handlers = []
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
    email_pass = []
    email_fail = []

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

    # PULL REPORT DATA FROM SQL SERVER
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
            # print(row.RPT_EXPRT_FILE_TYP,row.RPT_SRVR_LOC_DESC,row.RPT_FILE_NM, filepath)
            get_rpt(row.RPT_EXPRT_FILE_TYP,row.RPT_SRVR_LOC_DESC,row.RPT_FILE_NM, filepath)
            rpt_pass.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM])
            logger.info("\tSUCCESS! Report ID: {0} \t| Name: {1}".format(row.RPT_EXPRT_DIM_ID,row.RPT_NM))
        except:
            rpt_fail.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM])
            logger.warning("\tFAIL! Report ID: {0} \t| Name: {1}".format(row.RPT_EXPRT_DIM_ID,row.RPT_NM))
    
    # LOG REPORT OUTPUT
    pass_rpts = len(rpt_pass)
    fail_rpts = len(rpt_fail)
    total_rpts = pass_rpts + fail_rpts

    # pass reports
    logger.info("{0} of {1} reports downloaded successfully.".format(pass_rpts,total_rpts))
    # fail reports
    logger.info("{0} of {1} reports failed to download.".format(fail_rpts,total_rpts))
    logger.info("Downloads complete")
    cursor.close()
#---------------------------------------------------------------------------------------------------------------
    rpt_list = []
    rpt_ids = ''

    #prepare email list
    for rpt in rpt_pass:
        rpt_list.append(rpt[0])
    
    rpt_ids = ','.join(str(x) for x in rpt_list)
    logger.info("Sucessful report ids: {0}".format(rpt_ids))
    
    cursor = cnxn.cursor()
    sql = """
            SELECT
            A.RPT_EXPRT_DIM_ID
            ,B.RPT_EXPRT_EMAIL_DIM_ID
            ,B.USER_EMAIL
            ,A.[RPT_NM]
            ,[RPT_FILE_NM] = A.[RPT_FILE_NM]
                + CAST(year(getdate()) as varchar(4))
                +'-'+case when len(month(getdate())) = 1 THEN '0' + cast(month(getdate()) as varchar(2)) ELSE CAST(month(getdate()) as varchar(2)) END
                +'-' +case when len(day(getdate())) = 1 THEN '0' + cast(day(getdate()) as varchar(2))  ELSE CAST(day(getdate()) as varchar(2)) END 
                + [RPT_EXPRT_FILE_TYP]
            FROM [ICG_WORK].[work].[RPT_EXPRT_DIM] A
            LEFT JOIN work.RPT_EXPRT_EMAIL_DIM B ON B.RPT_EXPRT_DIM_ID = A.RPT_EXPRT_DIM_ID
            WHERE 1=1
            AND A.ACTV_IND = 1
            AND A.EXPIR_DT IS NULL
            
            AND B.EXPIR_DT IS NULL
            AND A.RPT_EXPRT_DIM_ID IN (""" + rpt_ids + """)
        """
        # AND B.USER_EMAIL = 'gavin@sineanalytics.com'
    # print(sql)
    cursor.execute(sql)
                    
    logger.info("SQL email cursor executed")
    
    logger.info("Begin email routine")
    # Email REPORTS
    for row in cursor:
        rptnm = 'Attached Report: ' + row.RPT_NM
        body = """
                <p>""" + row.RPT_NM + """ report is attached.</p>
                <p>Please reply to this message with any questions.</p>
                <p>Thank you,</p>
                <p>Sine Analytics</p>
               """
        try:
            email(row.USER_EMAIL, rptnm, body, filename=row.RPT_FILE_NM, filepath=filepath,html="yes")
            email_pass.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM,row.USER_EMAIL,row.RPT_EXPRT_EMAIL_DIM_ID])
            logger.info("\tSENT! Email ID: {0} \t| Report: {1} \t| To: {2}".format(row.RPT_EXPRT_EMAIL_DIM_ID,row.RPT_NM,row.USER_EMAIL))            
        except:
            email_fail.append([row.RPT_EXPRT_DIM_ID,row.RPT_NM,row.USER_EMAIL,row.RPT_EXPRT_EMAIL_DIM_ID])
            logger.warning("email failed: {0} to {1} ".format(row.RPT_NM,row.USER_EMAIL))
            # return False

    # LOG EMAIL OUTPUT
    pass_email = len(email_pass)
    fail_email = len(email_fail)
    total_email = pass_email + fail_email
    
    # pass emails
    logger.info("{0} of {1} emails were sent successfully.".format(pass_email,total_email))
    # fail emails
    logger.info("{0} of {1} emails failed to send.".format(fail_email,total_email))
    logger.info("emails complete")
    cursor.close()
    logger.info('Finished')
    email('gavin@sineanalytics.com',dt + ' email export log','Log file is attached.',filename=logfile,filepath=logpath)
    

def get_rpt(file_format, report_loc, file_name, filepath):
    # print('start get report')
    config = configparser.ConfigParser()
    config.read('config.ini')
    username = config['ssrs']['username']
    password = config['ssrs']['password']
    base_url = config['ssrs']['base_url']

    default_report_loc = r'%2fRevenue%20Reporting%2fBranch+Revenue+Report&BRANCH=222-Team Nunzio'
    
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

    if file_name:
        out_file = filepath + file_name + file_ext

    session = Session()
    session.auth = HttpNtlmAuth(username=username, password=password)
    response = session.get(url)
    
    try:
        with open(out_file, 'wb') as f:
            f.write(response.content)
    except:
        print('failed to write to file')

    session.close()
    
if __name__ == "__main__":
    main()