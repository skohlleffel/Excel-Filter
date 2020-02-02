from django.shortcuts import render, redirect
# Create your views here.
from django.core.files import File
from .models import Excel
from .forms import FileForm
# Create your views here.

def file_list(request):
    context = {
        'Excel': Excel.objects.all()
               }
    return render(request, 'file_list.html', context)

def upload_file(request):
    if request.method == 'POST':
        uploaded_file = request.FILES['xlsx']
        import pandas as pd
        name_check = []
        numeric_names = []
        del_comp_list = []
        del_email_list = []
        deleted_companies = {}
        deleted_emails = {}
        deleted_items = []
        df = pd.read_excel(uploaded_file)
        # converts data in cells to string
        """In order to work: 'Job Title' needs to be 'Title', 'Email Address' needs to be 'Email', 'Company Name' needs to be 'Company', if name is given as first and last in 2 seperate columns needs to be titled as 'First Name' and 'Last Name', if full name given in one column needs to be 'Full Name'"""
        # drops rows with nan values in email and name columns
        df.dropna(subset=['Email', 'First Name'], inplace=True)
        # replaces nan values with blanks
        df.fillna('', inplace=True)
        # converts data in sheet to string values
        df = df.applymap(str)
        # list of unwanted companies
        partners_remove = ['EVIRDIS', 'AGILEDSS', 'AMEEXUSA', 'D-TC', 'FMR', 'FINICITY', 'PWC', 'RACKSPACE',
                           'STRSOFTWARE', 'FIVETRAN', 'GOOGLE CLOUD', 'GOOGLE CLOUD PLATFORM', 'INFORMATICA',
                           'LOOKER',
                           'MATILLION', 'MICROSOFT', 'SIGMA', 'TABLEAU', 'TALEND', 'ALTERYX', 'ATTUNITY', 'CHARTIO',
                           'DATAGUISE', 'DOMO', 'MICROSTRATEGY', 'QUBOLE', 'SEGMENT', 'STITCH', 'WHERESCAPE',
                           'ABINITIO', 'AGINITY', 'ALATION', 'ALEX SOLUTIONS', 'ASCEND', 'ATSCALE', 'AVORA',
                           'BIRST',
                           'BRYTEFLOW', 'CDATA', 'CLOVERETL', 'COLLIBRA', 'COMPILERWORKS', 'DATABRICKS', 'DATAIKU',
                           'DATAROBOT', 'DATA ROBOT', 'DIYOTTA', 'DOMINO DATA LAB', 'ETLEAP', 'GOODDATA',
                           'GOOD DATA',
                           'H20.AI', 'HEAP ANALYTICS', 'HEVO DATA', 'HUNTERS.AI', 'HVR', 'IBM COGNOS', 'IMMUTA',
                           'INFORMATION BUILDERS', 'IRI', 'KEBOOLA', 'LACEWORK', 'LYFTRON', 'MPARTICLE',
                           'NEXT PATHWAY',
                           'PAXATA', 'PERISCOPE DATA', 'RIVERY', 'SALESFORCE', 'EINSTEIN', 'SECUPI',
                           'SESAME SOFTWARE',
                           'SISENSE', 'SNAPLOGIC', 'SNOWPLOW', 'SQLSTREAM', 'STREAMSETS', 'STRIIM', 'SYNCSORT',
                           'TAMR',
                           'THOUGHTSPOT', 'VISION.BI', 'WANDISCO', 'WINNOW', 'XPLENTY', 'ZETARIS', 'ZOOMDATA',
                           'BMC',
                           'AWS', 'AMAZON', 'AMAZON WEB SERVICES', 'ABILIS', 'ACCENTURE', 'ACTINVISION', 'ADVECTAS',
                           'AE BUSINESS SOLUTIONS', 'AGILIZ', 'AGINIC', 'ALPHAZETTA', 'ALTIS', 'ANALYTICS8',
                           'APPEX',
                           'APTITIVE', 'ARCHETYPE', 'ARETO', 'ARKATECHTURE', 'AWARESERVICES', 'AWH', 'AXIS GROUP',
                           'B.TELLIGENT', 'BIG DATA DIMENSION', 'BIG DATA SOLUTIONS', 'BILLIGENCE', 'BIMANU',
                           'BITFACTOR', 'BIYOND', 'BIZDATA', 'BIZONE', 'BRILLAR', 'BUSINESS & DECISION',
                           'BYTECODE IO',
                           'CAPGEMINI', 'CASERTA', 'CERTUS SOLUTIONS', 'CERULIUM', 'CERVELLO', 'CIS CONSULTING',
                           'CLARITY INSIGHTS', 'CLOUDNILE', 'CLOUDTEN', 'CLOUDWICK', 'CODER CO', 'COGNIZANT',
                           'CONTINO',
                           'CORE COMPETE', 'CRIMSON MACAW', 'CRITICALMINDS', 'DATABRIGHT', 'DATAFACTZ', 'DATALYTYX',
                           'DATAMEANING', 'DATAROBOT', 'DATASTICIANS', 'DECISIONMINDS', 'DECISIVE DATA', 'DELOITTE',
                           'DEPT', 'DEVOTEAM', 'DIGITAL MANAGEMENT, INC', 'DMI', 'DIGITAL MANAGEMENT INC',
                           'DUNN SOLUTIONS', 'ELIZA', 'EPAM', 'EULIDIA', 'FAIRWAY TECHNOLOGIES', 'FIRN ANALYTICS',
                           'FOREST GROVE', 'FRESH GRAVITY', 'G2O', 'GENSQUARED', 'GRAYTRAILS', 'HALPENFIELD',
                           'HASHMAP',
                           'ICON INTEGRATION', 'IMPETUS', 'IN516HT', 'INFEENY', 'INFINITY WORKS', 'INFOCENTRIC',
                           'INFOCEPTS', 'INFOREADY', 'INFOSYS', 'INITIONS', 'INTELIA', 'INTERWORKS', 'INTICITY',
                           'IQVIA TECHNOLOGIES', 'IRONSIDE', 'KABEL', 'KADENZA', 'KEBOOLA', 'KETL', 'KEYRUS',
                           'KINAESIS', 'KMPG FRANCE', 'KNOWIT', 'LARSEN & TOUBRO INFOTECH',
                           'LARSEN AND TOUBRO INFOTECH', 'LAUNCH CONSULTING', 'LEADING EDGE IT', 'LEVATAS',
                           'LINCUBE',
                           'LOGIC', 'MECHANICAL ROCK', 'MICROSTRATEGIES', 'MIKAN', 'MILLERSOFT', 'MINORO',
                           'MOMENTUM CONSULTING', 'MOSER', 'NATIVEML', 'NEUDESIC', 'NOW CONSULTING',
                           'NOW CONSULTING (WHERESCAPE)', 'WHERESCAPE', 'NTT DATA', 'OBILLIGENCE',
                           'ONE SIX SOLUTIONS',
                           'ONEBRIDGE', 'OSS GROUP', 'PANDATA', 'PANDERA SYSTEMS', 'PASSIO CONSULTING', 'PDX',
                           'PERFORMANCE ARCHITECTS', 'PERSISTENT', 'PRECOCITY', 'QUANDATICS', 'QUANTYCA',
                           'QUINSCAPE',
                           'RCG GLOBAL', 'RED PILL ANALYTICS', 'REDKITE INTELLIGENCE', 'RXP', 'RXP SERVICES',
                           'SAAMA',
                           'SAGGEZZA', 'SATALYST', 'SDG GROUP', 'SERVIAN', 'SHERPA CONSULTING', 'SIMPLE MACHINES',
                           'SIRIUS', 'SLALOM', 'SMART ASSOCIATES', 'SMARTRONIX', 'SOFTSERVE', 'SOLITA', 'SONATA',
                           'SONRA', 'SPARKHOUND', 'SPARKS', 'SUTTER MILLS', 'SYNERGY', 'TAIL WIND TECHNOLOGIES',
                           'TAMGROUP', 'TAYSOLS', 'TCS', 'TCS-GLOBAL', 'TECH MAHINDRA', 'TEKNION', 'TEKPARTNERS',
                           'TEKSYSTEMS', 'TENZING', 'THE ANALYTICS ACADEMY', 'TIMMARON GROUP', 'TRACE3', 'TREDENCE',
                           'TRIANZ', 'TROPOS', 'USEREADY', 'UST GLOBAL', 'VANTAGE DATA', 'VERSENT', 'VISION BI',
                           'VISUAL BI', 'WAVICLE DATA SOLUTIONS', 'WIPRO', 'WORKCENTIC', 'YSANCE']
        # list of unwanted email services
        email_remove = ['evirdis', 'agiledss', 'ameexusa', 'd-tc', 'fmr', 'finicity', 'pwc', 'rackspace',
                        'strsoftware',
                        'wnco', 'earthlink', 'att', 'bellsouth', 'gmail', 'yahoo', 'gmx', 'hotmail', 'fastmail',
                        'aol',
                        'zoho', 'trashmail', 'icloud', 'protonmail', 'outlook', 'msn']
        # uppercases all of the companies in the companies column
        df['Company'] = df['Company'].str.upper()
        # for sheets with 'Full Name' Column
        if 'Full Name' in df.columns:
            # Splits the Full Name column into two columns (First Name, Last Name)
            df['First Name'], df['Last Name'] = df['Full Name'].str.split(' ', 1).str
            # drops the full name column
            df = df.drop(columns=['Full Name'])
            # places the first and last name columns to be the first columns in the new sheet
            df = df[['Last Name'] + [col for col in df.columns if col != 'Last Name']]
            df = df[['First Name'] + [col for col in df.columns if col != 'First Name']]
        # Capitalizes the first and last names in the sheet
        df['First Name'] = df['First Name'].str.lower()
        df['First Name'] = df['First Name'].str.capitalize()
        df['Last Name'] = df['Last Name'].str.lower()
        df['Last Name'] = df['Last Name'].str.capitalize()
        # creates a list of companies with names containing only numerical values if any are present
        for i in df['Company']:
            name_check.append(i)
            if i.isnumeric() == True:
                numeric_names.append(i)
        # initializes for loop to look at the rows and columns of the sheet
        for index, row in df.iterrows():
            # for loop removes rows with unwanted emails
            for i in email_remove:
                if i in row['Email']:
                    try:
                        df.drop(index, inplace=True)
                        del_email_list.append(row['Email'])
                    except KeyError:
                        pass
            # for loop removes rows with unwanted companies
            for i in partners_remove:
                if i in row['Company'] and row['Company'] != 'SIRIUSXM':
                    try:
                        df.drop(index, inplace=True)
                        del_comp_list.append(row['Company'])
                    except KeyError:
                        pass
            # for loop removes rows with companies that have numerical names
            for i in numeric_names:
                if i in row['Company']:
                    try:
                        df.drop(index, inplace=True)
                    except KeyError:
                        pass
        # capitalizes companies and job titles
        deleted_emails['Emails'] = del_email_list
        deleted_companies['Companies'] = del_comp_list
        deleted_items.append(deleted_companies)
        deleted_items.append(deleted_emails)
        print(deleted_items)
        df['Company'] = df['Company'].str.lower()
        df['Company'] = df['Company'].str.capitalize()
        df['Title'] = df['Title'].str.lower()
        df['Title'] = df['Title'].str.title()
        df.fillna('x')
        # creates new .csv file with new data
        new_file = Excel(xlsx=File(df.to_csv(index=False)))
        new_file.save()
        return redirect('file_list')
    else:
        form = FileForm()
        return render(request, 'upload_file.html', {
            'form': form
        })