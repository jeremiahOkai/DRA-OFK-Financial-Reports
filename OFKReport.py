"""XML Generating Script

This script allows the user to generate xml files.

It is assumed that the first row of the spreadsheet is the
location of the columns and the first column is the location of the indexes.

This tool accepts only excel (.xls, .xlsx) files as inputs.

Python 3.5 and above is required for executing this script. Inaddition the script requires that
'pandas, xml, numpy, datetime', be installed within the Python
environment you are running this script in.


This file can also be imported as a module and contains the following
functions:

    * generateXML - write xml to a file.
    * validate_XML_OFKFiles - validate all generate xml against DNB schema

"""

import pandas as pd
import xml.etree.ElementTree as ET
import os
from datetime import datetime
import numpy as np
import glob
import lxml.etree as LT
import logging

# This code is used for logging errors
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# error level, module name and error message
formatter = logging.Formatter('%(asctime)s:%(message)s')
file_handler = logging.FileHandler('OFKReport.log')  # name of error log
# level for which error to be logged (Error). Anything other error will be ignored.
file_handler.setLevel(logging.ERROR)
# format in which error should be reported
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()  # logging error to the console
# format in which error should be reported
stream_handler.setFormatter(formatter)

logger.addHandler(file_handler)  # adding file_handler settings to the logger
# adding stream_handler settings to the logger
logger.addHandler(stream_handler)


# define Python user-defined exceptions
class Error(Exception):
    """Base class for other exceptions"""
    pass


class EmptyValueError(Error):
    """Raised when the input value is empty"""
    pass


class IntegerValueError(Error):
    """Raised when the input value is too is not an integer"""
    pass


class StringValueError(Error):
    """Raised when the input value is not a string"""
    pass


class LengthValueRequiredError(Error):
    """Raised when the input value is not equal to the required length of values. eg. 'BE, NL' """
    pass


def generateXML(files):
    """Gets and generate the xml file.

    Parameters
    ----------
    filename : str
        The file location of the spreadsheet.

    Attributes
    ----------
    counter : int
        Used the keep track of the location of the spreadsheet.

    existing_subForms : list, empty
        Used to keep track of all OFK form tags which are in spreadsheet.

    dataframes : list, empty
        Used as a container to store form tag, subform and the dataframe.

    generated_file_date : datetime
        Used to keep track of the datetime when the form is generated.

    subformRegeltag : dict
        Used as a container to store subform and their corresponding control tag.

    formName : str
        Used to keep track of the form tag of the current worksheet

    subformName : str
        Used to keep track of the subform tag of the current worksheet

    output : str
        Output name of the file.

    columns_to_validate : dict, empty
        Used to track all column tags for which values are required. Used in the Error handling.

    Raises
    -------
    EmptyValueError
        if no input value is provided

    StringValueError
        if a value not equal to string is provided

    IntegerValueError
        if a vlaue not equal to integer is provided

    AssertionError
        if non-negative value is provided

    """
    subformRegeltag = {'AD-A': 'AlgDeeln', 'AD-C': 'DeelnAct', 'ADO-C': 'OnrGoed', 'AEB-A': 'Aandelen',
                       'AEB-AI': 'Aandelen', 'AEBB-A': 'Aandelen', 'AEBB-AI': 'Aandelen', 'AEBB-G': 'GeldmarktPap',
                       'AEBB-K': 'KapitaalmarktPap', 'AEBB-KGI': 'Schuldpapier', 'AEB-G': 'GeldmarktPap', 'AEB-K': 'KapitaalmarktPap',
                       'AEB-KGI': 'Schuldpapier', 'AEI-A': 'Aandelen', 'AEI-AI': 'Aandelen', 'AEI-G': 'GeldmarktPap', 'AEI-K': 'KapitaalmarktPap',
                       'AEI-KGI': 'Schuldpapier', 'AEL-A': 'Aandelen', 'AEL-AI': 'Aandelen', 'AEL-G': 'GeldmarktPap', 'AEL-K': 'KapitaalmarktPap',
                       'AEL-KGI': 'Schuldpapier', 'AEN-A': 'Aandelen', 'AEN-AI': 'Aandelen', 'AENB-A': 'Aandelen', 'AENB-AI': 'Aandelen',
                       'AENB-G': 'GeldmarktPap', 'AENB-K': 'KapitaalmarktPap', 'AENB-KGI': 'Schuldpapier', 'AENL-A': 'Aandelen',
                       'AENL-AI': 'Aandelen', 'AENL-G': 'GeldmarktPap', 'AENL-K': 'KapitaalmarktPap', 'AENL-KGI': 'Schuldpapier',
                       'AEN-G': 'GeldmarktPap', 'AEN-K': 'KapitaalmarktPap', 'AEN-KGI': 'Schuldpapier', 'AEU-A': 'Aandelen',
                       'AEU-AI': 'Aandelen', 'AEU-G': 'GeldmarktPap', 'AEU-K': 'KapitaalmarktPap', 'AEU-KGI': 'Schuldpapier',
                       'ANF-C': 'ActivaNietFin', 'ANF-CGM': 'ActivaNietFin', 'ANF-CGJ': 'ActivaNietFin', 'AO-DI': 'DeelnIntInst',
                       'AOE-A': 'Aandelen', 'AOE-AI': 'Aandelen', 'AOE-G': 'GeldmarktPap', 'AOE-K': 'KapitaalmarktPap',
                       'AOE-KGI': 'Schuldpapier', 'AO-FL': 'LeasesUG', 'AO-HK': 'HandUGK', 'AO-HL': 'HandUGL',
                       'AO-HY': 'HypoUG', 'AO-LK': 'LeningUGK', 'AO-LL': 'LeningUGL', 'AO-OK': 'OverigeUGK',
                       'AO-OL': 'OverigeUGL', 'AO-RC': 'RecCourant', 'AO-RP': 'RepoUG', 'AR': 'StichKap', 'AV-LP': 'TechVoorz',
                       'AV-VV': 'LopenAanspr', 'BENB-A': 'Aandelen', 'BENB-AI': 'Aandelen', 'BENB-G': 'GeldmarktPap',
                       'BENB-K': 'KapitaalmarktPap', 'BENB-KGI': 'Schuldpapier', 'BT': 'BalansTotaal', 'D-FB': 'Futures',
                       'D-FN': 'Futures', 'DO-FB': 'Futures', 'D-OK': 'OptiesGekocht', 'DO-OK': 'OptiesGekocht', 'DO-OS': 'OptiesGeschr',
                       'DO-OTR': 'OTCDerivaten', 'DO-OTV': 'OTVDerivaten', 'D-OS': 'OptiesGeschr', 'D-OTR': 'OTCDerivaten', 'D-OTV': 'OTVDerivaten',
                       'IO-GO': 'OntwHulpGebond', 'IO-OO': 'OntwHulpSchenk', 'IO-XH': 'InkomOverdracht', 'GD-ECM': 'DienstExtraConcern',
                       'GD-ICM': 'DienstIntraConcern', 'GD-GLM': 'RandLGebruiksLicent', 'GD-RLM': 'RandLReprodLicent', 'GD-ECJ': 'DienstExtraConcern',
                       'GD-ICJ': 'DienstIntraConcern', 'GD-GLJ': 'RandLGebruiksLicent', 'GD-RLJ': 'RandLReprodLicent', 'IWB': 'IntrWaarde',
                       'KO-KW': 'SchuldKwijtSch', 'KO-OG': 'OnrGoedBtlOv', 'KO-SR': 'Stamrechten', 'KO-VO': 'OverKapitOverdr',
                       'PD-A': 'AlgDeeln', 'PD-C': 'DeelnPass', 'PEN-A': 'Aandelen', 'PEN-AI': 'Aandelen', 'PENB-A': 'Aandelen',
                       'PENB-AI': 'Aandelen', 'PENB-G': 'GeldmarktPap', 'PENB-K': 'KapitaalmarktPap', 'PENB-KGI': 'Schuldpapier',
                       'PENL-A': 'Aandelen', 'PENL-AI': 'Aandelen', 'PENL-G': 'GeldmarktPap', 'PENL-K': 'KapitaalmarktPap',
                       'PENL-KGI': 'Schuldpapier', 'PEN-G': 'GeldmarktPap', 'PEN-K': 'KapitaalmarktPap', 'PEN-KGI': 'Schuldpapier',
                       'PN-OS': 'NedTegenpartij', 'PO-FL': 'LeasesOG', 'PO-HK': 'HandOGK', 'PO-HL': 'HandOGL', 'PO-LK': 'LeningOGK',
                       'PO-LL': 'LeningOGL', 'PO-OK': 'OverigeOGK', 'PO-OL': 'OverigeOGL', 'PO-RP': 'RepoOG', 'PV-LP': 'TechVoorz',
                       'PV-OV': 'OverVoorz', 'PV-VV': 'LopenAanspr', 'SB-K': 'ParticipatieUGK', 'SB-L': 'ParticipatieUGL',
                       'SN-K': 'ParticipatieOGK', 'SN-L': 'ParticipatieOGL', 'WE-A': 'Aandelen', 'WE-AI': 'Aandelen', 'WE-G': 'GeldmarktPap',
                       'WE-K': 'KapitaalmarktPap', 'WE-KGI': 'Schuldpapier', 'WI-A': 'Aandelen', 'WI-AI': 'Aandelen', 'WI-G': 'GeldmarktPap',
                       'WI-K': 'KapitaalmarktPap', 'WI-KGI': 'Schuldpapier', 'WVA-B': 'Bestemming', 'WVA-R': 'Resulaten', 'WVA-Z': 'Resulaten',
                       'WVB-B': 'Baten', 'WVB-L': 'Bedrijfskosten', 'WVB-O': 'Bedrijfskosten', 'WVB-S': 'Loonkosten', 'WVP': 'WinstVerlPremUitk',
                       'WVP-ZA': 'AanvZorg', 'WVP-ZZ': 'ZorgZvw', 'WVT-BL': 'TotBatenLasten', 'WVU-B': 'WinstVerlBuiten', 'WVU-L': 'LatenUitz'}

    for file in files:
        logger.debug(f'Generating xml for --> {file}\n')
        dataframes, existing_subForms = [], []
        columns_to_validate = {}

        # XML Header Information
        generated_file_date = datetime.today().strftime('%Y-') + '0' + \
            str(int(datetime.today().strftime('%m')) -
                1)  # change 2 to 1 and put a condition here
        root = ET.Element('OFK-K')
        root.set("xmlns", "bb.dnb.nl")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        xsi_value = "bb.dnb.nl " + 'OFK-K.' + generated_file_date + '.xsd'
        root.set('xsi:schemaLocation', xsi_value)

        # XML Data information
        generated_file_date = datetime.today().strftime('%Y-') + '0' + \
            str(int(datetime.today().strftime('%m')) -
                1)  # generate datetime information eg.2020-04
        rappOpmerkingen = ET.SubElement(root, 'rappOpmerkingen')
        rappOpmerkingen.text = 'OFK ' + generated_file_date

        output = os.path.splitext(file)[0]
        test = pd.ExcelFile(file)  # open Excelfile
        for sheet in test.sheet_names:  # Loop through all sheets names
            # Select Formulierenoverzicht worksheet.
            if sheet == 'Formulierenoverzicht':
                continue
            else:
                # Get each subform profile and process data information.
                df = test.parse(sheet)

                formName = df.reset_index(drop=True).iloc[0].values.tolist()[
                    0]  # get form tag from current worksheet

                df.index = pd.Series(df.index).replace(
                    np.nan, 'No label')  # set null indexes as 'no label'

                # select column values with index name kolomtag
                cols = df.loc['Kolomtag'].tolist()

                df.columns = cols

                cols = [x for x in cols if str(x) != 'nan']

                df = df[cols]

                if sheet in ['AD-A', 'PD-A']:

                    # drop indexes of dataframe
                    df.drop(index=['No label', 'Kolomtag'], inplace=True)

                    df.columns = cols  # replace header of dataframe with kolomtag column values

                else:
                    # drop indexes and first column of dataframe
                    df.drop(index=['No label', 'Kolomtag'],
                            columns=[cols[0]], inplace=True)

                    df.reset_index(inplace=True)  # reset index of dataframe

                    df.columns = cols  # replace header of dataframe with kolomtag column values

                # replace all null fields with empty string
                df = df.replace(np.nan, '', regex=True)

                if sheet not in columns_to_validate:  # dict with sheet as key and column headers as values

                    columns_to_validate[sheet] = cols

                # the code below drops columns automatically for all worksheet which DNB has explicitly blocked for no values to be populated.
                if sheet in ['AD-C', 'PD-C']:
                    df.drop(df.columns[-1], axis=1, inplace=True)
                elif sheet in ['AEB-A', 'AEN-A']:
                    df.drop(
                        df.columns[[1, 11, 12, 13, 14, 15, 16]], axis=1, inplace=True)
                elif sheet == 'AEB-AI':
                    df.drop(
                        df.columns[[2, 3, 7, 8, 11, 12, 13, 14, 15, 16]], axis=1, inplace=True)
                elif sheet in ['AEB-G', "AEB-K", 'AEN-G', 'AEN-K']:
                    df.drop(df.columns[[1, 17]], axis=1, inplace=True)
                elif sheet == 'AEB-KGI':
                    df.drop(
                        df.columns[[2, 3, 7, 8, 11, 12, 15, 16, 17]], axis=1, inplace=True)
                elif sheet == 'AEN-AI':
                    df.drop(df.columns[[3, 7, 8, 11, 12, 13,
                                        14, 15, 16]], axis=1, inplace=True)
                elif sheet == 'AEN-KGI':
                    df.drop(
                        df.columns[[3, 7, 8, 11, 12, 15, 16, 17]], axis=1, inplace=True)
                elif sheet in ['AO-FL', 'AO-HL', 'AO-LK', 'AO-LL', 'AO-RP']:
                    df.drop(df.columns[[13]], axis=1, inplace=True)
                elif sheet == 'AO-HY':
                    df.drop(df.columns[[2, 13]], axis=1, inplace=True)
                elif sheet in ['AO-OK', 'AO-OL']:
                    df.drop(df.columns[[10, 11, 13, 14, 15]],
                            axis=1, inplace=True)
                elif sheet == 'AO-RC':
                    df.drop(df.columns[[10, 11, 14, 15]], axis=1, inplace=True)
                elif sheet == 'D-FB':
                    df.drop(df.columns[[1, 3, 4, 7, 8, 9, 10]],
                            axis=1, inplace=True)
                elif sheet in ['D-OK', 'D-OS']:
                    df.drop(df.columns[[1, 7, 8]], axis=1, inplace=True)
                elif sheet in ['D-OTR', 'D-OTV']:
                    df.drop(df.columns[[1, 8, 9]], axis=1, inplace=True)
                elif sheet == 'PEN-A':
                    df.drop(
                        df.columns[[1, 10, 11, 12, 13, 14, 15]], axis=1, inplace=True)
                elif sheet == 'PEN-AI':
                    df.drop(
                        df.columns[[6, 7, 10, 11, 12, 13, 14, 15]], axis=1, inplace=True)
                elif sheet == 'PEN-KGI':
                    df.drop(
                        df.columns[[6, 7, 10, 11, 14, 15, 16]], axis=1, inplace=True)
                elif sheet in ['PEN-G', 'PEN-K']:
                    df.drop(df.columns[[1, 16]], axis=1, inplace=True)
                elif sheet in ['PO-OK', 'PO-OL']:
                    df.drop(df.columns[[10, 11, 12, 14, 15]],
                            axis=1, inplace=True)
                elif sheet == 'PV-OV':
                    df.drop(df.columns[[1, 2]], axis=1, inplace=True)
                elif sheet in ['PO-FL', 'PO-HL', 'PO-LK', 'PO-LL', 'PO-RP']:
                    df.drop(df.columns[[12]], axis=1, inplace=True)
                elif sheet == 'WVA-B':
                    df.drop(df.columns[[1]], axis=1, inplace=True)
                elif sheet in ['WVB-B', 'WVB-L', 'WVB-S']:
                    df.drop(df.columns[[1, 3]], axis=1, inplace=True)

                if sheet == 'AO-RC':
                    print(df.head())

                df = df.T  # transpose dataframe
                # append form, worksheet and transposed dataframe.
                dataframes.append((formName, sheet, df))

        # Loop through list and verify fields within the worksheet before creating the xml fields below.
        try:
            for data in dataframes:  # loop through list
                # loop through transposed dataframe (dict)
                for key, values in data[2].items():
                    for k, v in values.items():
                        # check condition to see if sheet is in the range, if yes, verify if all field values are satisfied
                        if data[1] in ['AD-C', 'PD-C']:
                            if k.strip() in columns_to_validate[data[1]][3:-1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:6] + columns_to_validate[data[1]][7:8]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['AD-A', 'PD-A']:
                            if k.strip() in columns_to_validate[data[1]][3:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                            elif k.strip() in columns_to_validate[data[1]][1:3] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'ADO-C':
                            if k.strip() in columns_to_validate[data[1]][2:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][2:5] + columns_to_validate[data[1]][-2:-1]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['AEB-A', 'AEN-A']:
                            if k.strip() in columns_to_validate[data[1]][4:11] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][2:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'AEB-AI':
                            if k.strip() in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][9:11] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError

                        elif data[1] in ['AEB-G', "AEB-K", 'AEN-G', 'AEN-K']:
                            if k.strip() in columns_to_validate[data[1]][4:-1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7] + columns_to_validate[data[1]][13:15]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][2:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'AEB-KGI':
                            if k.strip() in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][9:11] + columns_to_validate[data[1]][13:15]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7] + columns_to_validate[data[1]][13:15]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError

                        elif data[1] == 'AEN-AI':
                            if k.strip() in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][9:11] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError
#
                        elif data[1] == 'AEN-KGI':
                            if k.strip() in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][9:11] + columns_to_validate[data[1]][13:15]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7] + columns_to_validate[data[1]][13:15]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'ANF-C':
                            if k.strip() in columns_to_validate[data[1]][1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][2:4]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"

                        elif data[1] in ['AO-FL', 'AO-HL', 'AO-LK', 'AO-LL', 'AO-RP']:
                            if k.strip() in columns_to_validate[data[1]][3:13] + columns_to_validate[data[1]][-2:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:13] + columns_to_validate[data[1]][-1:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'AO-HY':
                            if k.strip() in columns_to_validate[data[1]][3:13] + columns_to_validate[data[1]][-2:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:13] + columns_to_validate[data[1]][-1:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['AO-OK', 'AO-OL']:
                            if k.strip() in columns_to_validate[data[1]][3:10] + columns_to_validate[data[1]][12:13]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:10] + columns_to_validate[data[1]][12:13]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'AO-RC':
                            if k.strip() in columns_to_validate[data[1]][3:10] + columns_to_validate[data[1]][12:14]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:6] + columns_to_validate[data[1]][12:14]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'D-FB':
                            if k.strip() in columns_to_validate[data[1]][5:7]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][5:7]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['D-OK', 'D-OS']:
                            if k.strip() in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][-2:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:7] + columns_to_validate[data[1]][-2:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][2:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['D-OTR', 'D-OTV']:
                            if k.strip() in columns_to_validate[data[1]][4:8] + columns_to_validate[data[1]][-3:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:8] + columns_to_validate[data[1]][-2:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][2:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['GD-ECM', 'GD-ICM']:
                            if k.strip() in columns_to_validate[data[1]][2:4]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][2:4]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:2] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'PEN-A':
                            if k.strip() in columns_to_validate[data[1]][3:10] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:6] + columns_to_validate[data[1]][9:10] + columns_to_validate[data[1]][-1:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'PEN-AI':
                            if k.strip() in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][8:10] + columns_to_validate[data[1]][-1:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:10] + columns_to_validate[data[1]][-1:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'PEN-KGI':
                            if k.strip() in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][8:10] + columns_to_validate[data[1]][13:15]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:10] + columns_to_validate[data[1]][13:15]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['PEN-G', 'PEN-K']:
                            if k.strip() in columns_to_validate[data[1]][3:-1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:14] + columns_to_validate[data[1]][-2:-1]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] in ['PO-OK', 'PO-OL']:
                            if k.strip() in columns_to_validate[data[1]][3:10] + columns_to_validate[data[1]][-3:-2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:10] + columns_to_validate[data[1]][-3:-2]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'PV-OV':
                            if k.strip() in columns_to_validate[data[1]][3:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][4:6]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"

                        elif data[1] in ['PO-FL', 'PO-HL', 'PO-LK', 'PO-LL', 'PO-RP']:
                            if k.strip() in columns_to_validate[data[1]][3:12] + columns_to_validate[data[1]][-3:]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k in columns_to_validate[data[1]][3:6] + columns_to_validate[data[1]][9:12] + columns_to_validate[data[1]][-3:-2] + columns_to_validate[data[1]][-1:]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() in columns_to_validate[data[1]][1:3]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError
                                if k.strip() == 'Land' and len(v) != 2:
                                    raise LengthValueRequiredError

                        elif data[1] == 'WVA-B':
                            if k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError

                        elif data[1] in ['WVB-B', 'WVB-L', 'WVB-S']:
                            if k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k == columns_to_validate[data[1]][2]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"

                        elif data[1] in ['WVU-B', 'WVU-L']:
                            if k.strip() == columns_to_validate[data[1]][2]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError
                                if k == columns_to_validate[data[1]][2]:
                                    assert(
                                        int(v) >= 0), f"Non-negative value required at {k,data[1]}"
                            elif k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != str:
                                    raise StringValueError

                        elif data[1] == 'WVA-R':
                            if k.strip() == columns_to_validate[data[1]][1]:
                                if v == '':
                                    raise EmptyValueError
                                if type(v) != int:
                                    raise IntegerValueError

        except IntegerValueError:
            logger.error(f'Error Ocurred in {file}!!!')
            logger.error(
                f'Integer value is required in Column "{k}" of worksheet "{data[1]}"')
        except AssertionError as e:
            logger.error('Error Ocurred in {file}!!!')
            logger.error(e)
        except EmptyValueError:
            logger.error(f'Error Ocurred in {file}!!!')
            logger.error(
                f'Column "{k}" value of worksheet "{data[1]}" cannot be empty')
        except StringValueError:
            logger.error(f'Error Ocurred in {file}!!!')
            logger.error(
                f'String value is required in Column "{k}" of worksheet "{data[1]}"')
        except LengthValueRequiredError:
            logger.error('Error Ocurred in {file}!!!')
            logger.error(
                f'Example of accept values required column "Land" are "BE, NL, AZ" etc.')
        else:
            for data in dataframes:
                if data[0] in existing_subForms:
                    subformName = data[1]
                    subformName = ET.SubElement(formName, subformName)
                    for key, values in data[2].items():
                        control_tag = subformRegeltag[data[1]]
                        control_tag = ET.SubElement(subformName, control_tag)
                        for k, v in values.items():
                            k = ET.SubElement(control_tag, k)
                            k.text = str(v).strip()
                else:
                    formName = data[0]
                    formName = ET.SubElement(root, formName)
                    subformName = data[1]
                    subformName = ET.SubElement(formName, subformName)
                    for key, values in data[2].items():
                        control_tag = subformRegeltag[data[1]]
                        control_tag = ET.SubElement(subformName, control_tag)
                        for k, v in values.items():
                            k = ET.SubElement(control_tag, k)
                            k.text = str(v).strip()
                    existing_subForms.append(data[0])

            tree = ET.ElementTree(root)  # end of xml tree
            tree.write(output + '.xml', encoding="UTF-8",
                       xml_declaration=True)  # write xml to file
            logger.debug(f'Done generating xml for {file}\n')


def validate_XML_OFKFiles(files):
    for file in files:
        xml_file = LT.parse(file)
        xml_validator = LT.XMLSchema(file="OFK-K.2020-03.xsd")
        is_valid = xml_validator.validate(xml_file)
        if is_valid:
            logger.debug(f'{file} has successfully been validated!')
        else:
            logger.error(f'Validation for {file} was unsuccessful!')


if __name__ == '__main__':
    path = ''  # specify path or location to excel files
    os.chdir(path)
    files = []
    types = ('*.xls', '*.xlsx')
    for ext in types:
        files.extend(glob.glob(ext))
    generateXML(files)
    xmls = glob.glob("*.xml")
    validate_XML_OFKFiles(xmls)
