"""XML Generating Script

This script allows the user to generate xml files.

It is assumed that the first row of the spreadsheet is the
location of the columns and the first column is the location of the indexes.

This tool accepts only excel (.xls, .xlsx) files as inputs.

Python 3.5 and above is required for executing this script. In addition the script requires that
'pandas, xml, numpy, datetime', be installed within the Python
environment you are running this script in.


This file can also be imported as a module and contains the following
functions:

    * convert_to_xml - write xml to a file.
    * validate_XML_AIFM - validate all generate xml against DNB schema

"""

import pandas as pd
import xml.etree.ElementTree as ET
import os
from datetime import datetime
import numpy as np
import glob
import logging
import lxml.etree as LT
import sys

# define Python user-defined exceptions


class Error(Exception):
    """Base class for other exceptions"""
    pass


class EmptyValueError(Error):
    """Raised when the input value is empty"""
    pass


class DomainValueError(Error):
    """Raised when the input value is does not corresponds with the domain values required"""
    pass


class LengthValueRequiredError(Error):
    """Raised when the input value exceeds the max length required for that field """
    pass


class NoFilesFoundError(Error):
    """Raised when the input value exceeds the max length required for that field """
    pass


class UnassinedIntegerError(Error):
    """Raised when the input value is not an UnassignedInteger """
    pass


""" This code is used for logging errors """
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


def convert_to_xml(files):
    """Get excel files and convert to XML.

    Parameters
    ----------
    files : str
        The file location of the spreadsheet.

    output : str
        Output name of the file.


    Raises
    -------
    EmptyValueError
        if no input value is provided

    DomainValueError
        if input value is does not corresponds with the domain values required

    LengthValueRequiredError
        if value exceed the length of what is required

    NoFilesFoundError
        if no files are selected.

    UnassinedIntegerError
        if value is not an unsigned integer (not a negative number or contain decimal).


    """

    try:
        if len(files) == 0:
            raise NoFilesFoundError(
                "No excel files selected. Specify path to excel document and try again!")

        for file in files:
            logger.debug(f'Generating xml for --> {file}\n')
            output_file = os.path.splitext(file)[0]
            df = pd.read_excel(file, header=None)
            df = df.replace(np.nan, '', regex=True)
            df.columns = ['xmlTags', 'Id',
                          'XMLDescription', 'Input_1', 'Input_2']
            df.xmlTags = df.xmlTags.str.strip('<>')

            headerRows = [str(i) for i in range(1, 4)]
            headerFileKeys = df[df.Id.isin(headerRows)][[
                'xmlTags', 'Input_1']].values.tolist()

            if headerFileKeys[0][1] and headerFileKeys[1][1] == "":
                raise EmptyValueError(
                    f" {headerFileKeys[0][0]} and {headerFileKeys[1][0]} fields cannot be empty! ")

            generated_on = datetime.today().strftime('%Y-%m-%dT%H:%M:%S')
            root = ET.Element('AIFMReportingInfo')
            root.set("xsi:noNamespaceSchemaLocation", "AIFMD_DATMAN_V1.2.xsd")
            root.set('CreationDateAndTime', generated_on)
            root.set(headerFileKeys[0][0], str(headerFileKeys[0][1]).strip())
            root.set(headerFileKeys[1][0], str(headerFileKeys[1][1]).strip())
            root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

            sectionRows = [str(i) for i in range(4, 10)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}

            AIFMRecordInfo = ET.SubElement(root, 'AIFMRecordInfo')
            for k, v in headerSectionKeys.items():
                if v != "":
                    if k not in ['ReportingPeriodStartDate', 'ReportingPeriodEndDate']:
                        k = ET.SubElement(AIFMRecordInfo, k)
                        k.text = str(v).strip()
                    else:
                        k = ET.SubElement(AIFMRecordInfo, k)
                        k.text = str(v.date()).strip()
                else:
                    raise EmptyValueError(f"{k} field cannot be empty!")

            sectionRows = [str(i) for i in range(10, 16)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}
#
            if headerSectionKeys['AIFMReportingObligationChangeFrequencyCode'] != "":
                AIFMReportingObligationChangeFrequencyCode = ET.SubElement(
                    AIFMRecordInfo, 'AIFMReportingObligationChangeFrequencyCode')
                AIFMReportingObligationChangeFrequencyCode.text = str(
                    headerSectionKeys['AIFMReportingObligationChangeFrequencyCode']).strip()

            if headerSectionKeys['AIFMReportingObligationChangeContentsCode'] != "":
                AIFMReportingObligationChangeContentsCode = ET.SubElement(
                    AIFMRecordInfo, 'AIFMReportingObligationChangeContentsCode')
                AIFMReportingObligationChangeContentsCode.text = str(
                    headerSectionKeys['AIFMReportingObligationChangeContentsCode']).strip()

            if headerSectionKeys['AIFMReportingObligationChangeFrequencyCode'] or headerSectionKeys['AIFMReportingObligationChangeContentsCode'] != "":
                if headerSectionKeys['AIFMReportingObligationChangeQuarter'] != "":
                    AIFMReportingObligationChangeQuarter = ET.SubElement(
                        AIFMRecordInfo, 'AIFMReportingObligationChangeQuarter')
                    AIFMReportingObligationChangeQuarter.text = str(
                        headerSectionKeys['AIFMReportingObligationChangeQuarter']).strip()
                else:
                    raise EmptyValueError(
                        "AIFMReportingObligationChangeQuarter field cannot be empty!")

            if headerSectionKeys['LastReportingFlag'] != "":
                LastReportingFlag = ET.SubElement(
                    AIFMRecordInfo, 'LastReportingFlag')
                LastReportingFlag.text = str(
                    headerSectionKeys['LastReportingFlag']).lower().strip()
            else:
                raise EmptyValueError(
                    "LastReportingFlag field cannot be empty!")

            if headerSectionKeys['QuestionNumber'] and headerSectionKeys['AssumptionDescription'] == "":
                if len(headerSectionKeys['AssumptionDescription']) > 300:
                    raise LengthValueRequiredError(
                        f'AssumptionDescription string required in this field should not be greater 300!')

                QuestionNumber = ET.SubElement(
                    AIFMRecordInfo, 'QuestionNumber')
                QuestionNumber.text = str(
                    headerSectionKeys['QuestionNumber']).strip()
                AssumptionDescription = ET.SubElement(
                    AIFMRecordInfo, 'AssumptionDescription')
                AssumptionDescription.text = str(
                    headerSectionKeys['AssumptionDescription']).strip()

            sectionRows = [str(i) for i in range(16, 22)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}
            for k, v in headerSectionKeys.items():
                if v != "":
                    k = ET.SubElement(AIFMRecordInfo, k)
                    k.text = str(v).strip()
                else:
                    raise EmptyValueError(f"{k} field cannot be empty!")

            identifierRows = [str(i) for i in range(22, 26)]
            AIMFIdentifiers = df[df.Id.isin(identifierRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            AIMFIdentifiers = {i[0]: i[1] for i in AIMFIdentifiers}

            AIFMCompleteDescription = ET.SubElement(
                AIFMRecordInfo, 'AIFMCompleteDescription')

            if AIMFIdentifiers:
                AIFMIdentifier = ET.SubElement(
                    AIFMCompleteDescription, 'AIFMIdentifier')

                if AIMFIdentifiers['AIFMIdentifierLEI'] != "":
                    AIFMIdentifierLEI = ET.SubElement(
                        AIFMIdentifier, 'AIFMIdentifierLEI')
                    AIFMIdentifierLEI.text = str(
                        AIMFIdentifiers['AIFMIdentifierLEI']).strip()

                if AIMFIdentifiers['AIFMIdentifierBIC'] != "":
                    AIFMIdentifierBIC = ET.SubElement(
                        AIFMIdentifier, 'AIFMIdentifierBIC')
                    AIFMIdentifierBIC.text = str(
                        AIMFIdentifiers['AIFMIdentifierBIC']).strip()

                if AIMFIdentifiers['ReportingMemberState'] and AIMFIdentifiers['ReportingMemberState'] != "":
                    ReportingMemberState = ET.SubElement(
                        AIFMIdentifier, 'ReportingMemberState')
                    ReportingMemberState.text = str(
                        AIMFIdentifiers['ReportingMemberState']).strip()
                    AIFMNationalCode = ET.SubElement(
                        AIFMIdentifier, 'AIFMNationalCode')
                    AIFMNationalCode.text = str(
                        AIMFIdentifiers['AIFMNationalCode']).strip()

            ranks = [i for i in range(1, 6)]
            principalMRows = ['1st', '2nd', '3rd', '4th', '5th']
            principalMarkets = df[df.Id.isin(principalMRows)][[
                'XMLDescription', 'Input_1', 'Input_2']].values.tolist()

            counter = 0
            AIFMPrincipalMarkets = ET.SubElement(
                AIFMCompleteDescription, 'AIFMPrincipalMarkets')

            for principalMarket in principalMarkets:
                if principalMarket[0] in ['MIC', 'XXX', 'OTC', 'NOT']:
                    AIFMFivePrincipalMarket = ET.SubElement(
                        AIFMPrincipalMarkets, 'AIFMFivePrincipalMarket')
                    Ranking = ET.SubElement(
                        AIFMFivePrincipalMarket, 'Ranking')
                    Ranking.text = str(ranks[counter]).strip()
                    MarketIdentification = ET.SubElement(
                        AIFMFivePrincipalMarket, 'MarketIdentification')
                    MarketCodeType = ET.SubElement(
                        MarketIdentification, 'MarketCodeType')
                    MarketCodeType.text = str(principalMarket[0]).strip()

                    if principalMarket[0] == "MIC" and principalMarket[1] == "":
                        raise EmptyValueError(
                            "MarketCode is required for MIC market codes!")
#
                    if principalMarket[0] == "MIC" and len(principalMarket[1]) > 4:
                        raise LengthValueRequiredError(
                            f'Maximum length of 4 is required for MarketCode!')

                    if principalMarket[0] == "MIC":
                        MarketCode = ET.SubElement(
                            MarketIdentification, 'MarketCode')
                        MarketCode.text = str(principalMarket[1]).strip()
#
                    if principalMarket[0] != "NOT":
                        if type(principalMarket[2]) != int or principalMarket[2] < 0:
                            raise UnassinedIntegerError(
                                "AggregatedValueAmount must be not contain decimals & should not be a negative number. Check value in row 32-36 column D")
                        else:
                            AggregatedValueAmount = ET.SubElement(
                                AIFMFivePrincipalMarket, 'AggregatedValueAmount')
                            AggregatedValueAmount.text = str(
                                principalMarket[2]).strip()

                    counter += 1

                else:
                    raise DomainValueError(
                        "Required principal market values in for AIFM trades are 'MIC', 'XXX', 'OTC' & 'NOT'. Check values in rows 32-36, column C of template")
#

            principalIRow = [i for i in range(1, 6)]
            principalInstruments = df[df.Id.isin(principalIRow)][[
                'Id', 'XMLDescription', 'Input_1']].values.tolist()

            AIFMPrincipalInstruments = ET.SubElement(
                AIFMCompleteDescription, 'AIFMPrincipalInstruments')

            for principalInstrument in principalInstruments:
                if principalInstrument[1] == "":
                    raise EmptyValueError(
                        f"SubAssetType field cannot be empty! Check values in rows 40-44, column C of template")

                if principalInstrument[2] == "":
                    raise EmptyValueError(
                        f"AggregatedValueAmount field cannot be empty! Check values in rows 40-44, column C of template")

                AIFMPrincipalInstrument = ET.SubElement(
                    AIFMPrincipalInstruments, 'AIFMPrincipalInstrument')
                Ranking = ET.SubElement(AIFMPrincipalInstrument, 'Ranking')
                Ranking.text = str(principalInstrument[0]).strip()
                SubAssetType = ET.SubElement(
                    AIFMPrincipalInstrument, 'SubAssetType')
                SubAssetType.text = str(principalInstrument[1]).strip()

                if principalInstrument[1] != 'NTA_NTA_NOTA':
                    if type(principalInstrument[2]) != int or principalInstrument[2] < 0:
                        raise UnassinedIntegerError(
                            "AggregatedValueAmount must be UnassigneIinteger (not contain decimals & not negative). Check value in row 40-44 column D")
                    else:
                        AggregatedValueAmount = ET.SubElement(
                            AIFMPrincipalInstrument, 'AggregatedValueAmount')
                        AggregatedValueAmount.text = str(
                            principalInstrument[2]).strip()

            valuesRows = (str(i) for i in range(33, 39))
            principalValues = df[df.Id.isin(
                valuesRows)][['xmlTags', 'Input_1']].values.tolist()
            principalValues = {i[0]: i[1] for i in principalValues}

            if principalValues['AUMAmountInEuro'] != "":
                if type(principalValues['AUMAmountInEuro']) != int or principalValues['AUMAmountInEuro'] < 0:
                    raise UnassinedIntegerError(
                        "AUMAmountInEuro must be UnassigneIinteger (not contain decimals and not negative). Check value in row 46 column D")

                AUMAmountInEuro = ET.SubElement(
                    AIFMCompleteDescription, 'AUMAmountInEuro')
                AUMAmountInEuro.text = str(
                    principalValues['AUMAmountInEuro']).strip()
            else:
                raise EmptyValueError(
                    f"AUMAmountInEuro field cannot be empty! Check value in row 46 column D")

            AIFMBaseCurrencyDescription = ET.SubElement(
                AIFMCompleteDescription, 'AIFMBaseCurrencyDescription')

            if principalValues['AUMAmountInBaseCurrency'] and principalValues['BaseCurrency'] != "":
                if type(principalValues['AUMAmountInBaseCurrency']) != int or principalValues['AUMAmountInBaseCurrency'] < 0:
                    raise UnassinedIntegerError(
                        "AUMAmountInBaseCurrency must be UnassigneIinteger (not contain decimals & not negative). Check value in row 47 column D")

                BaseCurrency = ET.SubElement(
                    AIFMBaseCurrencyDescription, 'BaseCurrency')
                BaseCurrency.text = str(
                    principalValues['BaseCurrency']).upper().strip()

                AUMAmountInBaseCurrency = ET.SubElement(
                    AIFMBaseCurrencyDescription, 'AUMAmountInBaseCurrency')
                AUMAmountInBaseCurrency.text = str(
                    principalValues['AUMAmountInBaseCurrency']).strip()

                if str(principalValues['BaseCurrency']).upper().strip() != "EUR":
                    if principalValues['FXEURReferenceRateType'] and principalValues['FXEURRate'] != "":
                        FXEURReferenceRateType = ET.SubElement(
                            AIFMBaseCurrencyDescription, 'FXEURReferenceRateType')
                        FXEURReferenceRateType.text = str(
                            principalValues['FXEURReferenceRateType']).strip()
                        FXEURRate = ET.SubElement(
                            AIFMBaseCurrencyDescription, 'FXEURRate')
                        FXEURRate.text = str(
                            principalValues['FXEURRate']).strip()
                    else:
                        raise EmptyValueError(
                            f"FXEURReferenceRateType or FXEURRate fields cannot be empty! Check value in rows 49-50 column D")

                if str(principalValues['BaseCurrency']).upper().strip() or principalValues['FXEUROtherReferenceRateDescription'] == "OTH":
                    FXEUROtherReferenceRateDescription = ET.SubElement(
                        AIFMBaseCurrencyDescription, 'FXEUROtherReferenceRateDescription')
                    FXEUROtherReferenceRateDescription.text = str(
                        principalValues['FXEUROtherReferenceRateDescription']).strip()

            tree = ET.ElementTree(root)
            tree.write(output_file + '.xml',
                       encoding="UTF-8", xml_declaration=True)
            logger.debug(f'Done generating xml for {file}\n')

    except Exception as e:
        logger.error(e)


def validate_XML_AIFM(files):
    for file in files:
        xml_file = LT.parse(file)
        xml_validator = LT.XMLSchema(file="AIFMD_DATMAN_V1.2.xsd")
        is_valid = xml_validator.validate(xml_file)
        if is_valid:
            logger.debug(f'{file} has successfully been validated!')
        else:
            logger.error(f'Validation for {file} was unsuccessful!')


if __name__ == '__main__':
    path = 'Enter path for the file'
    os.chdir(path)
    files = []
    types = ('*.xls', '*.xlsx')
    for ext in types:
        files.extend(glob.glob(ext))
    convert_to_xml(files)
    xmls = glob.glob("*.xml")
    validate_XML_AIFM(xmls)
    sys.exit(0)
