""" Python configuration file """

LOG_LEVEL = 'INFO'
LOG_FORMAT = '<g>{time:YYYY-MM-DD HH:mm:ss.SSS}</g> | ' \
             '<level>{level: <8}</level> - <level>{message}</level>'
LOG_FILE = './logs/country_financial_reports_{time}.log'

INPUT_DIR = 'input/'
TEMPLATE_DIR = 'template/'
OUTPUT_DIR = 'output/'

MISSING_FILE_MESSAGE = '{} is not found in the input directory'
EXISTENT_FILE_MESSAGE = '{} is found in the input directory'
MISSING_OUTPUT_FILE_MESSAGE = '{} is not found in the output directory. Run inputReport.py'
EXISTENT_OUTPUT_FILE_MESSAGE = '{} is found in the output directory'
INVALID_MAPPING_MESSAGE = 'Output cell mapping in row {} is not valid. ' \
                          'Please check the mapping file.'
INVALID_ALIAS_MESSAGE = 'Alias "{}" is not valid. Check entire row {} in file {}.'
MAPPING_ERROR_MESSAGE = 'Error in evaluating statement "{}", check row {} in file {}'
MISSING_VALUE_ERROR = 'Value required for statement "{}" is missing, check row {} in file {}'
INCORRECT_VALUE_ERROR = 'Data type of value required for statement "{}" is incorrect, ' \
                        'check row {} in file {}'
ALIAS_NOT_FOUND_MESSAGE = 'Can not find, omit {}, {}'

INPUT_COUNTRY_FILE = '{} Country Financials_{}.xlsx'
COUNTRY_CAPITAL_FILE = 'Country Capital file_{}.xlsx'
WEEKLY_COUNTRY_FILE = 'Weekly country financials working file_cob {}.xlsx'
OUTPUT_FILE_FORMAT = '{} Country Financials_{}_TopCoder.xlsx'
TEMPLATE_FORMAT = 'Country Financials_template.xlsx'

ALIAS_FILE_INPUT = 'mapping/alias/alias_input.csv'
ALIAS_FILE_PB = 'mapping/alias/alias_pb.csv'
ALIAS_FILE_SEA_REV = 'mapping/alias/alias_sea_rev.csv'
ALIAS_FILE_SEA_EXP = 'mapping/alias/alias_sea_exp.csv'
ALIAS_FILE_SEA_PTI = 'mapping/alias/alias_sea_pti.csv'
ALIAS_FILE_SEA_COUNTRY = 'mapping/alias/alias_{}.csv'
ALIAS_FILE_SEA_COUNTRY_TREND = 'mapping/alias/alias_{}_trend.csv'
ALIAS_FILE_AM = 'mapping/alias/alias_am.csv'
ALIAS_FILE_AM_ESSBASE = 'mapping/alias/alias_am_essbase.csv'
ALIAS_FILE_AFG = 'mapping/alias/alias_afg.csv'
ALIAS_FILE_MKTS = 'mapping/alias/alias_mkts.csv'
ALIAS_FILE_IBCM = 'mapping/alias/alias_ibcm.csv'
ALIAS_FILE_APO = 'mapping/alias/alias_apo.csv'
ALIAS_FILE_WMCO = 'mapping/alias/alias_wmco.csv'
ALIAS_FILE_EXP = 'mapping/alias/alias_exp.csv'
ALIAS_FILE_WEEKLY = 'mapping/alias/alias_weekly.csv'
ALIAS_FILE_WEEKLY_PREV = 'mapping/alias/alias_weekly_prev.csv'

MAPPING_INPUT_TAB = 'mapping/rules/MappingInputTab.csv'
MAPPING_PB_TAB = 'mapping/rules/MappingPB{}Tab.csv'
MAPPING_SEA_REV_TAB = 'mapping/rules/MappingSeaRevTab.csv'
MAPPING_SEA_EXP_TAB = 'mapping/rules/MappingSeaExpTab.csv'
MAPPING_SEA_PTI_TAB = 'mapping/rules/MappingSeaPtiTab.csv'
MAPPING_SEA_COUNTRY_TAB = 'mapping/rules/mapping_{}_tab.csv'
MAPPING_SEA_COUNTRY_TREND_TAB = 'mapping/rules/mapping_{}_trend_tab.csv'
MAPPING_AM_TAB = 'mapping/rules/mapping_{}_tab.csv'
MAPPING_AFG_TAB = 'mapping/rules/MappingAFGTab.csv'
MAPPING_IBCM_TAB = 'mapping/rules/MappingIBCMTab.csv'
MAPPING_MKTS_TAB = 'mapping/rules/MappingMktsTab.csv'
MAPPING_APO_TAB = 'mapping/rules/mapping_apo_tab.csv'
MAPPING_WMCO_TAB = 'mapping/rules/mapping_wmco_tab.csv'
MAPPING_EXP_TAB = 'mapping/rules/mapping_exp_{}_tab.csv'

# For Grouping column exists
AFG_GROUP_EXISTED_COUNTRIES = [
    'India, and Market Group, Indian Sub-Continent',
]

MKTS_GROUP_EXISTED_COUNTRIES = [
    'Australia',
    'India, and Market Group, Indian Sub-Continent',
    'Japan',
    'Singapore',
]

MKTS_HALF_GROUP_EXISTED_COUNTRIES = [
    'India',
]

AFG_VS_BUD_NOT_EXISTED_COUNTRIES = [
    'Australia',
    'Greater China',
    'India',
    'SEA & FM',
    'Singapore',
]

AFG_VS_BUD_HALF_NOT_EXISTED_COUNTRIES = [
    'Japan',
]

MKTS_VS_BUD_EXISTED_COUNTRIES = [
    'Australia',
    'Japan',
    'Singapore',
]

MKTS_EQ_MANAGEMENT_EXISTED_COUNTRIES = [
    'Malaysia',
    'Philippines',
    'Thailand',
    'Singapore',
    'Frontier Markets',
    'Indonesia',
]

MKTS_NOTES_TEXT = '- FY19 country FP as of 3 May 2019 based on APAC Feb BoD submission.'

IBCM_NO_ECM_QTD_COUNTRIES = [
    'Malaysia',
    'Frontier Markets',
    'Indonesia',
    'Korea',
    'Philippines',
    'Thailand',
    'Vietnam',
]

IBCM_GROUP_EXISTED_COUNTRIES = [
    'Australia',
    'India, and Market Group, Indian Sub-Continent',
    'Japan',
    'Singapore',
]

# Memo: AFG Revenues (excl Excess Funding) row does not exist
AFG_MEMO_NOT_EXISTED_COUNTRIES = [
    'Thailand',
    'Indonesia',
]
COUNTRIES_LIST = ['Australia', 'China', 'India', 'Japan', 'Korea', 'SEA&FM', \
                  'Frontier Markets', 'Indonesia', 'Malaysia', 'Philippines', \
                  'Singapore', 'Thailand', 'Vietnam']

COUNTRY_NAMES = {'Australia' : 'Australia',
                 'India'     : 'India, and Market Group, Indian Sub-Continent',
                 'Singapore' : 'Singapore',
                 'Indonesia' : 'Indonesia',
                 'Malaysia'  : 'Malaysia',
                 'Thailand'  : 'Thailand',
                 'Philippines': 'Philippines',
                 'Frontier Markets' : 'Frontier Markets',
                 'Vietnam'   : 'Vietnam',
                 'SEA&FM' : 'SEA&FM',
                 'China' : 'Gr China',
                 'Japan' : 'Japan',
                 'Korea' : 'Korea'
                }

GROUP1_COUNTRIES = ['Australia', 'Japan', 'Singapore']
GROUP2_COUNTRIES = ['Gr China', 'India, and Market Group, Indian Sub-Continent', 'SEA&FM']
OTHER_COUNTRIES = ['Korea', 'Frontier Markets', 'Indonesia', 'Malaysia', 'Philippines', \
                   'Thailand', 'Vietnam']

PB_INPUT_TAB = {'Gr China': ('PB Legacy GC', 'PB GC Others'),
                'India, and Market Group, Indian Sub-Continent': ("PB (Onshore)", "PB (NRI)"),
                'SEA&FM': ('PB Legacy SEA', 'PB SEA Others')
                }

SEA_COUNTRY_ATTR = {
    'rev': ('SEA country Revenue (2)', 'sea_rev'),
    'exp': ('SEA country Expenses', 'sea_exp'),
    'pti': ('SEA country PTI', 'sea_pti')
    }

SEA_COUNTRY_TREND_ATTR = {
    'rev': ('SEA country revenue trends', 'SEA country Revenue (2)', 'sea_rev_trend'),
    'exp': ('SEA country Expense trend', 'SEA country Expenses', 'sea_exp_trend'),
    'pti': ('SEA country PTI trend', 'SEA country PTI', 'sea_pti_trend')
    }

COUNTRY_DATE = '' # to be updated by program
NON_PB_CODE = {'Australia' : 'O.P_AN',
               'India'     : 'O.P_SA_IND',
               'Singapore' : 'O.P_SA_SGP',
               'Indonesia' : 'O.P_SA_INO',
               'Malaysia'  : 'O.P_SA_MAL',
               'Thailand'  : 'O.P_SA_THI',
               'Philippines': 'O.P_SA_PHI',
               'Frontier Markets' : 'O.P_NJ_OA',
               'Vietnam'   : 'O.P_SA_VIE',
               'SEA & FM' : 'O.P_NJ_SEAFM',
               'Greater China' : 'O.P_NJ_GCN',
               'Japan' : 'O.P_JP',
               'Korea' : 'O.P_NJ_KOR'
              }
PB_MARKET_CODE = {'Australia' : '69911_AGGR',
                  'India'     : '63620_AGGR',
                  'Singapore' : '60637_AGGR',
                  'Indonesia' : '60284_AGGR',
                  'Malaysia'  : '60577_AGGR',
                  'Thailand'  : '60008_AGGR',
                  'Philippines': '60089_AGGR',
                  'Vietnam'   : '0',
                  'Frontier Markets' : '0',
                  'SEA & FM' : '60912_AGGR',
                  'Greater China' : '60958_AGGR',
                  'Japan' : '63690_AGGR',
                  'Korea' : '0'
                 }

PB_MARKET_DESC = {'Australia' : 'CS Australia & EAM AU (TF) (69911)',
                  'India'     : 'CS India (TF) (63620)',
                  'Singapore' : 'Singapore Market (TF) (60637)',
                  'Indonesia' : 'Indonesia Market (TF) (60284)',
                  'Malaysia'  : 'Malaysia Market (TF) (60577)',
                  'Thailand'  : 'Thailand (TF) (60008)',
                  'Philippines': 'Philippines Market (TF) (60089)',
                  'Vietnam'   : '0',
                  'Frontier Markets' : '0',
                  'SEA & FM' : 'South East Asia (TF) (60912)',
                  'Greater China' : 'Greater China (TF) (60958)',
                  'Japan' : 'CS Japan (TF) (63690)',
                  'Korea' : '0'
                 }
