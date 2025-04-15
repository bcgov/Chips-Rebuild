import requests
import os
#import ssl
import csv
from datetime import datetime
import pandas as pd
import logging
from dotenv import load_dotenv
from ldap3 import Server, Connection, ALL
from dataclasses import fields
from epdObject import EDPObject

#email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def setup_logging():
    log_dir = "log"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        filename=os.path.join(log_dir, 'edp.log')
    )

logger = logging.getLogger(__name__)

class ActiveDirectoryQuery:
    def __init__(self, server_url, bind_dn, bind_password, base_dn):
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.server = Server(server_url, get_info=ALL)
        self.conn = Connection(self.server, bind_dn, bind_password, auto_bind=True)
        self.base_dn = base_dn
    
    noIDIR = 0
    totalProcessed = 0

    def returnNoIDIR(self):
        return self.noIDIR
    
    def returnTotalProcessed(self):
        return self.totalProcessed
    
    def query_employee(self, emp_id, work_Develop_Region, work_Regonal_District, work_Location_Group, 
                        psa_email, abc_Employee_Rule, PSA_Non_PSA_Rule, misc_App_Rule, active_Employee_Rule,
                        special_Cases_Rule, contact_Information_Flag, executive_Classification_Flag, work_address, flag3,
                        Action, Action_Date, Action_Descr, Action_Reason, Action_Reason_Descr, Age, Age_Group_2,
                        Base_EFFDT,Base_Function_CD,Base_Function_CD_Descr,Base_Grade,
                        Base_Jobcode,Base_Jobcode_Descr,Base_Position_NBR,Base_SAL_ADMIN_PLAN,Base_Step,BirthDate,Business_Title,BusinessUnit,BusinessUnitNumber,
                        ClassificationGroup,CmpnySeniorityAge,CmpnySeniorityDate,Department,DeptID,Duplicate,EFFDT,EFFSEQ,Email,
                        Empl_CTG,Empl_RCD,Empl_Status,EmplID,EmployeeCategoryShortDesc,EmployeeCategoryType,EmployeeStatusDesc,FirstName,FulltimeParttime_Calculated,
                        FulltimeParttime_JobData,Function_CD,FunctionDescription,GRADE,HireAge,HireDate,HR_Status,IncludedExcluded,Jobcode,Jobcode_Descr,
                        JobCodeDescGroup,LastName,LeaveAccrualDate,Level1,Level2,Level3,Level4,MinStartAge,MinStartDate,Name,NOC_Code,NOC_Code_Descr,
                        NOC_Code_EFF_STATUS,NOC_Code_EFF_STATUS_DESCR,NOC_Code_Occ_Group,OccupationalGroup,ORG,Organization,OrigHireDate,PersonalAddress1,PersonalAddress2,PersonalAddress3,
                        PersonalAddress4,PersonalCity,PersonalPostalCode,PersonalProv,Phone,Position_Descr,Position_NBR,PSA,RegStartAge,RegStartDate,Sal_Admin_Plan,
                        Sal_Admin_Plan_Decsr,Sector,SEQ,ServiceAge,ServiceDate,Sex,STD_Hours,STEP,Union_CD,Union_CD_Descr,WorkAddress1,WorkAddress2,WorkAddress3,WorkAddress4,WorkCity,
                        WorkDevelopmentRegion,WorkLocationGroup,WorkPostalCode,WorkProv,WorkRegionalDistrict):
        """
            Query the LDAP for details to fill it's portion of the report
        """
        

        try:
            #search_filter = f'(mail={Email})'
            search_filter = f'(employeeID={emp_id})'

            attributes = [
                'employeeID', 'sAMAccountName', 'mailboxOrgCode','postalCode', 'l', 'streetAddress', 'telephoneNumber',
                'bcgovGUID', 'mail'
            ]

            self.conn.search(
                search_base=self.base_dn,
                search_filter=search_filter,
                attributes=attributes
            )
            self.totalProcessed += 1
            if self.conn.entries:
                entry = self.conn.entries[0]
                return EDPObject(
                    # IDIR Fields
                    IDIR_EmplID=entry.employeeID.value if entry.employeeID else '',
                    IDIR_Userid=entry.sAMAccountName.value if entry.sAMAccountName else '',
                    IDIR_Mail_org=entry.mailboxOrgCode.value if entry.mailboxOrgCode else '',
                    idirPostal=entry.postalCode.value if entry.postalCode else '',
                    idirCity=entry.l.value if entry.l else '',
                    idirStreet=entry.streetAddress.value if entry.streetAddress else '',
                    idirPhone=entry.telephoneNumber.value if entry.telephoneNumber else '',
                    bcGUID=entry.bcgovGUID.value if entry.bcgovGUID else '',
                    idirEmail=entry.mail.value if entry.mail else '',
                    #PSA Report entries
                    personalRegion=work_Develop_Region,
                    regionalDistrict=work_Regonal_District,
                    economicalDevRegion=work_Location_Group,
                    ABCEmployeeRule= abc_Employee_Rule,
                    PSANonPSARule= PSA_Non_PSA_Rule,
                    miscAppRule=misc_App_Rule,
                    activeEmployeeRule=active_Employee_Rule,
                    specialCasesRule=special_Cases_Rule,
                    contactInformationFlag=contact_Information_Flag,
                    executiveClassificationFlag=executive_Classification_Flag,
                    # Flags
                    flagEmail= 'F' if (entry.mail.value == psa_email) else 'T',
                    MissmatchAddress = '' if (entry.streetAddress.value == work_address) else 'T', # Flags if Difference in address
                    InActiveDirectory='T', # Flags if no IDIR entry
                    MultipleRecordsAD= flag3,  # Flag if there are duplicate entries
                    # EDP API 
                    Action=Action,
                    Action_Date=Action_Date,
                    Action_Descr=Action_Descr,
                    Action_Reason=Action_Reason,
                    Action_Reason_Descr=Action_Reason_Descr,
                    Age=Age,
                    Age_Group_2=Age_Group_2,
                    Base_EFFDT=Base_EFFDT,
                    Base_Function_CD=Base_Function_CD,
                    Base_Function_CD_Descr=Base_Function_CD_Descr,
                    Base_Grade=Base_Grade,
                    Base_Jobcode=Base_Jobcode,
                    Base_Jobcode_Descr=Base_Jobcode_Descr,
                    Base_Position_NBR=Base_Position_NBR,
                    Base_SAL_ADMIN_PLAN=Base_SAL_ADMIN_PLAN,
                    Base_Step=Base_Step,
                    BirthDate=BirthDate,
                    Business_Title=Business_Title,
                    BusinessUnit=BusinessUnit,
                    businessUnitNumber=BusinessUnitNumber,
                    ClassificationGroup=ClassificationGroup,
                    CmpnySeniorityAge=CmpnySeniorityAge,
                    CmpnySeniorityDate=CmpnySeniorityDate,
                    Department=Department,
                    DeptID=DeptID,
                    Duplicate=Duplicate,
                    EFFDT=EFFDT,
                    EFFSEQ=EFFSEQ,
                    Email=Email,
                    Empl_CTG=Empl_CTG,
                    Empl_RCD=Empl_RCD,
                    Empl_Status=Empl_Status,
                    EmplID=EmplID,
                    EmployeeCategoryShortDesc=EmployeeCategoryShortDesc,
                    EmployeeCategoryType=EmployeeCategoryType,
                    EmployeeStatusDesc=EmployeeStatusDesc,
                    FirstName=FirstName,
                    FulltimeParttime_Calculated=FulltimeParttime_Calculated,
                    FulltimeParttime_JobData=FulltimeParttime_JobData,
                    Function_CD=Function_CD,
                    FunctionDescription=FunctionDescription,
                    GRADE=GRADE,
                    HireAge=HireAge,
                    HireDate=HireDate,
                    HR_Status=HR_Status,
                    IncludedExcluded=IncludedExcluded,
                    Jobcode=Jobcode,
                    Jobcode_Descr=Jobcode_Descr,
                    JobCodeDescGroup=JobCodeDescGroup,
                    LastName=LastName,
                    LeaveAccrualDate=LeaveAccrualDate,
                    Level1=Level1,
                    Level2=Level2,
                    Level3=Level3,
                    Level4=Level4,
                    MinStartAge=MinStartAge,
                    MinStartDate=MinStartDate,
                    Name=Name,
                    NOC_Code=NOC_Code,
                    NOC_Code_Descr=NOC_Code_Descr,
                    NOC_Code_EFF_STATUS=NOC_Code_EFF_STATUS,
                    NOC_Code_EFF_STATUS_DESCR=NOC_Code_EFF_STATUS_DESCR,
                    NOC_Code_Occ_Group=NOC_Code_Occ_Group,
                    OccupationalGroup=OccupationalGroup,
                    ORG=ORG,
                    Organization=Organization,
                    OrigHireDate=OrigHireDate,
                    PersonalAddress1=PersonalAddress1,
                    PersonalAddress2=PersonalAddress2,
                    PersonalAddress3=PersonalAddress3,
                    PersonalAddress4=PersonalAddress4,
                    PersonalCity=PersonalCity,
                    PersonalPostalCode=PersonalPostalCode,
                    PersonalProv=PersonalProv,
                    Phone=Phone,
                    Position_Descr=Position_Descr,
                    Position_NBR=Position_NBR,
                    PSA=PSA,
                    RegStartAge=RegStartAge,
                    RegStartDate=RegStartDate,
                    Sal_Admin_Plan=Sal_Admin_Plan,
                    Sal_Admin_Plan_Decsr=Sal_Admin_Plan_Decsr,
                    Sector=Sector,
                    SEQ=SEQ,
                    ServiceAge=ServiceAge,
                    ServiceDate=ServiceDate,
                    Sex=Sex,
                    STD_Hours=STD_Hours,
                    STEP=STEP,
                    Union_CD=Union_CD,
                    Union_CD_Descr=Union_CD_Descr,
                    WorkAddress1=WorkAddress1,
                    WorkAddress2=WorkAddress2,
                    WorkAddress3=WorkAddress3,
                    WorkAddress4=WorkAddress4,
                    WorkCity=WorkCity,
                    WorkDevelopmentRegion=WorkDevelopmentRegion,
                    WorkLocationGroup=WorkLocationGroup,
                    WorkPostalCode=WorkPostalCode,
                    WorkProv=WorkProv,
                    WorkRegionalDistrict=WorkRegionalDistrict
                )
            else:
                self.noIDIR += 1
                return EDPObject(
                    # IDIR Fields
                    IDIR_EmplID=emp_id,
                    IDIR_Userid='',
                    IDIR_Mail_org='',
                    idirPostal= '',
                    idirCity= '',
                    idirStreet= '',
                    idirPhone= '',
                    bcGUID= '',
                    idirEmail= '',
                    #PSA Report entries
                    personalRegion=work_Develop_Region,
                    regionalDistrict=work_Regonal_District,
                    economicalDevRegion=work_Location_Group,
                    ABCEmployeeRule= abc_Employee_Rule,
                    PSANonPSARule= PSA_Non_PSA_Rule,
                    miscAppRule=misc_App_Rule,
                    activeEmployeeRule=active_Employee_Rule,
                    specialCasesRule=special_Cases_Rule,
                    contactInformationFlag=contact_Information_Flag,
                    executiveClassificationFlag=executive_Classification_Flag,
                    # Flags
                    flagEmail= 'T',
                    MissmatchAddress = 'T', # Flags if Difference in address
                    InActiveDirectory='F', # Flags if no IDIR entry
                    MultipleRecordsAD= flag3,  # Flag if there are duplicate entries
                    # EDP API 
                    Action=Action,
                    Action_Date=Action_Date,
                    Action_Descr=Action_Descr,
                    Action_Reason=Action_Reason,
                    Action_Reason_Descr=Action_Reason_Descr,
                    Age=Age,
                    Age_Group_2=Age_Group_2,
                    Base_EFFDT=Base_EFFDT,
                    Base_Function_CD=Base_Function_CD,
                    Base_Function_CD_Descr=Base_Function_CD_Descr,
                    Base_Grade=Base_Grade,
                    Base_Jobcode=Base_Jobcode,
                    Base_Jobcode_Descr=Base_Jobcode_Descr,
                    Base_Position_NBR=Base_Position_NBR,
                    Base_SAL_ADMIN_PLAN=Base_SAL_ADMIN_PLAN,
                    Base_Step=Base_Step,
                    BirthDate=BirthDate,
                    Business_Title=Business_Title,
                    BusinessUnit=BusinessUnit,
                    businessUnitNumber=BusinessUnitNumber,
                    ClassificationGroup=ClassificationGroup,
                    CmpnySeniorityAge=CmpnySeniorityAge,
                    CmpnySeniorityDate=CmpnySeniorityDate,
                    Department=Department,
                    DeptID=DeptID,
                    Duplicate=Duplicate,
                    EFFDT=EFFDT,
                    EFFSEQ=EFFSEQ,
                    Email=Email,
                    Empl_CTG=Empl_CTG,
                    Empl_RCD=Empl_RCD,
                    Empl_Status=Empl_Status,
                    EmplID=EmplID,
                    EmployeeCategoryShortDesc=EmployeeCategoryShortDesc,
                    EmployeeCategoryType=EmployeeCategoryType,
                    EmployeeStatusDesc=EmployeeStatusDesc,
                    FirstName=FirstName,
                    FulltimeParttime_Calculated=FulltimeParttime_Calculated,
                    FulltimeParttime_JobData=FulltimeParttime_JobData,
                    Function_CD=Function_CD,
                    FunctionDescription=FunctionDescription,
                    GRADE=GRADE,
                    HireAge=HireAge,
                    HireDate=HireDate,
                    HR_Status=HR_Status,
                    IncludedExcluded=IncludedExcluded,
                    Jobcode=Jobcode,
                    Jobcode_Descr=Jobcode_Descr,
                    JobCodeDescGroup=JobCodeDescGroup,
                    LastName=LastName,
                    LeaveAccrualDate=LeaveAccrualDate,
                    Level1=Level1,
                    Level2=Level2,
                    Level3=Level3,
                    Level4=Level4,
                    MinStartAge=MinStartAge,
                    MinStartDate=MinStartDate,
                    Name=Name,
                    NOC_Code=NOC_Code,
                    NOC_Code_Descr=NOC_Code_Descr,
                    NOC_Code_EFF_STATUS=NOC_Code_EFF_STATUS,
                    NOC_Code_EFF_STATUS_DESCR=NOC_Code_EFF_STATUS_DESCR,
                    NOC_Code_Occ_Group=NOC_Code_Occ_Group,
                    OccupationalGroup=OccupationalGroup,
                    ORG=ORG,
                    Organization=Organization,
                    OrigHireDate=OrigHireDate,
                    PersonalAddress1=PersonalAddress1,
                    PersonalAddress2=PersonalAddress2,
                    PersonalAddress3=PersonalAddress3,
                    PersonalAddress4=PersonalAddress4,
                    PersonalCity=PersonalCity,
                    PersonalPostalCode=PersonalPostalCode,
                    PersonalProv=PersonalProv,
                    Phone=Phone,
                    Position_Descr=Position_Descr,
                    Position_NBR=Position_NBR,
                    PSA=PSA,
                    RegStartAge=RegStartAge,
                    RegStartDate=RegStartDate,
                    Sal_Admin_Plan=Sal_Admin_Plan,
                    Sal_Admin_Plan_Decsr=Sal_Admin_Plan_Decsr,
                    Sector=Sector,
                    SEQ=SEQ,
                    ServiceAge=ServiceAge,
                    ServiceDate=ServiceDate,
                    Sex=Sex,
                    STD_Hours=STD_Hours,
                    STEP=STEP,
                    Union_CD=Union_CD,
                    Union_CD_Descr=Union_CD_Descr,
                    WorkAddress1=WorkAddress1,
                    WorkAddress2=WorkAddress2,
                    WorkAddress3=WorkAddress3,
                    WorkAddress4=WorkAddress4,
                    WorkCity=WorkCity,
                    WorkDevelopmentRegion=WorkDevelopmentRegion,
                    WorkLocationGroup=WorkLocationGroup,
                    WorkPostalCode=WorkPostalCode,
                    WorkProv=WorkProv,
                    WorkRegionalDistrict=WorkRegionalDistrict
                )
        except Exception as e:
            self.logger.error(f"Error querying AD for EmpID {emp_id}: {e}")
            return None

def getReport (api_url, username, password, save_directory="./data"):
    try:
        # Make the API request with authentication
        response = requests.get(api_url, auth=(username, password))
        
        # Check if request was successful
        response.raise_for_status()
        
        # Create directory if it doesn't exist
        os.makedirs(save_directory, exist_ok=True)
        
        # Generate filename with format APIName_YYYYMMDD.csv
        today_date = datetime.now().strftime("%Y%m%d")
        filename = f"{'EDP'}_{today_date}.csv"
            
        # Full path for the file
        file_path = os.path.join(save_directory, filename)
        
        # Save the CSV data to file in binary mode to prevent line ending issues
        with open(file_path, 'wb') as file:
            file.write(response.content)
            
        logger.info(f"CSV successfully downloaded and saved to {file_path}")
        return file_path
        
    except requests.exceptions.RequestException as e:
        logger.error(f"Error downloading CSV: {e}")
        return None
    
def processReport(input_file, ad_query):
    """
    Process input CSV and query Active Directory for each employee
    
    :param input_file: Path to input CSV file
    :param ad_query: ActiveDirectoryQuery instance
    :return: List of EmployeeRecord objects
    """
    processed_records = []
    
    with open(input_file, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)

         

        for row in reader:
            emp_id = row.get('EmplID', '')
            work_Develop_Region = row.get('WorkDevelopmentRegion','')
            work_Regonal_District = row.get('WorkRegionalDistrict','')
            work_Location_Group = row.get('WorkLocationGroup','')
            psa_email = row.get('Email','')
            abc_Employee_Rule = row.get('Rule1', '')
            PSA_Non_PSA_Rule = row.get('Rule2','')
            misc_App_Rule = row.get('Rule3','')
            active_Employee_Rule = row.get('Rule4', '')
            special_Cases_Rule = row.get('Rule5','')
            contact_Information_Flag = row.get('Rule12','')
            executive_Classification_Flag = row.get('Rule13','')
            work_address = row.get('workWorkAddress1','')
            Action = row.get('Action', '')
            Action_Date = row.get('Action_Date', '')
            Action_Descr = row.get('Action_Descr', '')
            Action_Reason = row.get('Action_Reason', '')
            Action_Reason_Descr = row.get('Action_Reason_Descr', '')
            Age = row.get('Age', '')
            Age_Group_2 = row.get('Age_Group_2', '')
            Base_EFFDT = row.get('Base_EFFDT', '')
            Base_Function_CD = row.get('Base_Function_CD', '')
            Base_Function_CD_Descr = row.get('Base_Function_CD_Descr', '')
            Base_Grade = row.get('Base_Grade', '')
            Base_Jobcode = row.get('Base_Jobcode', '')
            Base_Jobcode_Descr = row.get('Base_Jobcode_Descr', '')
            Base_Position_NBR = row.get('Base_Position_NBR', '')
            Base_SAL_ADMIN_PLAN = row.get('Base_SAL_ADMIN_PLAN', '')
            Base_Step = row.get('Base_Step', '')
            BirthDate = row.get('BirthDate', '')
            Business_Title = row.get('Business_Title', '')
            BusinessUnitNumber = row.get('BU','')
            BusinessUnit = row.get('BusinessUnit', '')
            ClassificationGroup = row.get('ClassificationGroup', '')
            CmpnySeniorityAge = row.get('CmpnySeniorityAge', '')
            CmpnySeniorityDate = row.get('CmpnySeniorityDate', '')
            Department = row.get('Department', '')
            DeptID = row.get('DeptID', '')
            Duplicate = row.get('Duplicate', '')
            EFFDT = row.get('EFFDT', '')
            EFFSEQ = row.get('EFFSEQ', '')
            Email = row.get('Email', '')
            Empl_CTG = row.get('Empl_CTG', '')
            Empl_RCD = row.get('Empl_RCD', '')
            Empl_Status = row.get('Empl_Status', '')
            EmplID = row.get('EmplID', '')
            EmployeeCategoryShortDesc = row.get('EmployeeCategoryShortDesc', '')
            EmployeeCategoryType = row.get('EmployeeCategoryType', '')
            EmployeeStatusDesc = row.get('EmployeeStatusDesc', '')
            FirstName = row.get('FirstName', '')
            FulltimeParttime_Calculated = row.get('FulltimeParttime_Calculated', '')
            FulltimeParttime_JobData = row.get('FulltimeParttime_JobData', '')
            Function_CD = row.get('Function_CD', '')
            FunctionDescription = row.get('FunctionDescription', '')
            GRADE = row.get('GRADE', '')
            HireAge = row.get('HireAge', '')
            HireDate = row.get('HireDate', '')
            HR_Status = row.get('HR_Status', '')
            IncludedExcluded = row.get('IncludedExcluded', '')
            Jobcode = row.get('Jobcode', '')
            Jobcode_Descr = row.get('Jobcode_Descr', '')
            JobCodeDescGroup = row.get('JobCodeDescGroup', '')
            LastName = row.get('LastName', '')
            LeaveAccrualDate = row.get('LeaveAccrualDate', '')
            Level1 = row.get('Level1', '')
            Level2 = row.get('Level2', '')
            Level3 = row.get('Level3', '')
            Level4 = row.get('Level4', '')
            MinStartAge = row.get('MinStartAge', '')
            MinStartDate = row.get('MinStartDate', '')
            Name = row.get('Name', '')
            NOC_Code = row.get('NOC_Code', '')
            NOC_Code_Descr = row.get('NOC_Code_Descr', '')
            NOC_Code_EFF_STATUS = row.get('NOC_Code_EFF_STATUS', '')
            NOC_Code_EFF_STATUS_DESCR = row.get('NOC_Code_EFF_STATUS_DESCR', '')
            NOC_Code_Occ_Group = row.get('NOC_Code_Occ_Group', '')
            OccupationalGroup = row.get('OccupationalGroup', '')
            ORG = row.get('ORG', '')
            Organization = row.get('Organization', '')
            OrigHireDate = row.get('OrigHireDate', '')
            PersonalAddress1 = row.get('PersonalAddress1', '')
            PersonalAddress2 = row.get('PersonalAddress2', '')
            PersonalAddress3 = row.get('PersonalAddress3', '')
            PersonalAddress4 = row.get('PersonalAddress4', '')
            PersonalCity = row.get('PersonalCity', '')
            PersonalPostalCode = row.get('PersonalPostalCode', '')
            PersonalProv = row.get('PersonalProv', '')
            Phone = row.get('Phone', '')
            Position_Descr = row.get('Position_Descr', '')
            Position_NBR = row.get('Position_NBR', '')
            PSA = row.get('PSA', '')
            RegStartAge = row.get('RegStartAge', '')
            RegStartDate = row.get('RegStartDate', '')
            Sal_Admin_Plan = row.get('Sal_Admin_Plan', '')
            Sal_Admin_Plan_Decsr = row.get('Sal_Admin_Plan_Decsr', '')
            Sector = row.get('Sector', '')
            SEQ = row.get('SEQ', '')
            ServiceAge = row.get('ServiceAge', '')
            ServiceDate = row.get('ServiceDate', '')
            Sex = row.get('Sex', '')
            STD_Hours = row.get('STD_Hours', '')
            STEP = row.get('STEP', '')
            Union_CD = row.get('Union_CD', '')
            Union_CD_Descr = row.get('Union_CD_Descr', '')
            WorkAddress1 = row.get('WorkAddress1', '')
            WorkAddress2 = row.get('WorkAddress2', '')
            WorkAddress3 = row.get('WorkAddress3', '')
            WorkAddress4 = row.get('WorkAddress4', '')
            WorkCity = row.get('WorkCity', '')
            WorkDevelopmentRegion = row.get('WorkDevelopmentRegion', '')
            WorkLocationGroup = row.get('WorkLocationGroup', '')
            WorkPostalCode = row.get('WorkPostalCode', '')
            WorkProv = row.get('WorkProv', '')
            WorkRegionalDistrict = row.get('WorkRegionalDistrict', '')

            

            flag3 = 0

            ad_record = ad_query.query_employee(emp_id, work_Develop_Region, work_Regonal_District, work_Location_Group, 
                                                psa_email, abc_Employee_Rule, PSA_Non_PSA_Rule, misc_App_Rule, active_Employee_Rule,
                                                special_Cases_Rule, contact_Information_Flag, executive_Classification_Flag,
                                                work_address, flag3, Action, Action_Date, Action_Descr, Action_Reason, Action_Reason_Descr, Age, Age_Group_2,
                                                Base_EFFDT,Base_Function_CD,Base_Function_CD_Descr,Base_Grade,
                                                Base_Jobcode,Base_Jobcode_Descr,Base_Position_NBR,Base_SAL_ADMIN_PLAN,Base_Step,BirthDate,Business_Title,BusinessUnit,BusinessUnitNumber,
                                                ClassificationGroup,CmpnySeniorityAge,CmpnySeniorityDate,Department,DeptID,Duplicate,EFFDT,EFFSEQ,Email,
                                                Empl_CTG,Empl_RCD,Empl_Status,EmplID,EmployeeCategoryShortDesc,EmployeeCategoryType,EmployeeStatusDesc,FirstName,FulltimeParttime_Calculated,
                                                FulltimeParttime_JobData,Function_CD,FunctionDescription,GRADE,HireAge,HireDate,HR_Status,IncludedExcluded,Jobcode,Jobcode_Descr,
                                                JobCodeDescGroup,LastName,LeaveAccrualDate,Level1,Level2,Level3,Level4,MinStartAge,MinStartDate,Name,NOC_Code,NOC_Code_Descr,
                                                NOC_Code_EFF_STATUS,NOC_Code_EFF_STATUS_DESCR,NOC_Code_Occ_Group,OccupationalGroup,ORG,Organization,OrigHireDate,PersonalAddress1,PersonalAddress2,PersonalAddress3,
                                                PersonalAddress4,PersonalCity,PersonalPostalCode,PersonalProv,Phone,Position_Descr,Position_NBR,PSA,RegStartAge,RegStartDate,Sal_Admin_Plan,
                                                Sal_Admin_Plan_Decsr,Sector,SEQ,ServiceAge,ServiceDate,Sex,STD_Hours,STEP,Union_CD,Union_CD_Descr,WorkAddress1,WorkAddress2,WorkAddress3,WorkAddress4,WorkCity,
                                                WorkDevelopmentRegion,WorkLocationGroup,WorkPostalCode,WorkProv,WorkRegionalDistrict)
            
            if ad_record:
                processed_records.append(ad_record)

    #for record in processed_records:
     #   record.flag3 = 1 if checkDuplicates(input_file, record.idir) > 1 else 0
    
    processed_records = checkDuplicates(processed_records)

    return processed_records

def save_report(records, output_dir='EDP-reports'):
    """
    Save processed records to a new CSV report
    
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate filename with current date
    timestamp = datetime.now().strftime('%Y%m%d')
    output_file = os.path.join(output_dir, f'EPD_Report_{timestamp}.csv')
    
    # Get field names from EDPObject dataclass
    fieldnames = [field.name for field in fields(EDPObject)]
    
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for record in records:
            writer.writerow({field: getattr(record, field) for field in fieldnames})
    
    logger.info(f"Report saved: {output_file}")

def cleanAPIReport():
    try:
        path = "./data"
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)

            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    logger.info(f"{file_path} Deleted")
            
            except Exception as e:
                logger.error(f"Could not delete {file_path}. Error: {e}")
    except Exception as e:
        logger.error(f"An error occurred: {e}")

def checkDuplicates(report):
    logger.info("Checking Duplicates")
    seen_ids = {}
    duplicate_ids = set()

    for row in report:
        try:
            value = getattr(row, 'EmplID')
        except AttributeError:
            try:
                value = row['EmplID']
            except (TypeError, KeyError):
                logger.error(f"Warning: could not access EmplID on {row}")
                continue
        if value in seen_ids:
            duplicate_ids.add(value)
        else:
            seen_ids[value] = True
    
    for row in report:
        try:
            # First try attribute access
            value = getattr(row, 'EmplID')
            # Set the duplicate flag using setattr
            setattr(row, 'MultipleRecordsAD', value in duplicate_ids)
        except AttributeError:
            try:
                # Try dictionary-style access
                value = row['EmplID']
                # Set as dictionary item
                row['MultipleRecordsAD'] = value in duplicate_ids
            except (TypeError, KeyError):
                continue

    return report

def send_email(subject, message, from_email, to_emails, smtp_server):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ", ".join(to_emails)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server)
        server.sendmail(from_email, to_emails, msg.as_string())
        server.quit()
        logger.info(f'email sent sucessfully to {','.join(to_emails)}')
        return True
    except Exception as e:
        logger.error(f"Failed to send email: {str(e)}")
        return False



def main():
    # Start Logging
    setup_logging()
    #ENV
    load_dotenv()
    #Report Path
    psa_report_path = os.environ.get('psa_report_path')
    psa_report_user = os.environ.get('psa_report_user')
    psa_report_token = os.environ.get('psa_report_token')
    today_date = datetime.now().strftime("%Y%m%d")
    psa_filepath = f"./data/{'EDP'}_{today_date}.csv"

    #LDAP
    ad_server = os.environ.get('ldap_server_url')
    base_dn = os.environ.get('ldap_base_dn')
    ad_username = os.environ.get('ldap_username')
    ad_password = os.environ.get('ldap_password')

    #Email
    email_list = os.environ.get('email_list')
    error_emails = os.environ.get('error_emails')
    smtp_server = os.environ.get('smtp_server')
    from_email = os.environ.get('from_email')

    try:
        cleanAPIReport()

        logger.info('getting PSA report')
        getReport(psa_report_path, psa_report_user, psa_report_token)

        logger.info('PSA report acquired, querying IDIR')
        ad_query = ActiveDirectoryQuery(ad_server, ad_username, ad_password, base_dn)

        combinedReport =processReport(psa_filepath, ad_query)
        
        logger.info(f'Number of no IDIRS: {ad_query.returnNoIDIR()}')

        save_report(combinedReport)

        subject = "EDP Completed Succesfully {today_date}"
        msg = f"EDP completed. Total:{ad_query.returnTotalProcessed()}. Total no IDIR: {ad_query.returnNoIDIR()}"
        send_email(subject, msg, from_email, email_list, smtp_server)

    except Exception as e:
        subject = "EDP Error"
        msg = f"An error occurred: {e}"
        send_email(subject, msg, from_email, error_emails, smtp_server)
        logger.error(f"An error occurred: {e}")

    
    

if __name__ == "__main__":
    main()