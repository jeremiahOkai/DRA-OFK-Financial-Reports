#!/usr/bin/env python
# coding: utf-8


import pandas as pd
import xml.etree.ElementTree as ET
import os
from datetime import datetime
import numpy as np
import glob
import logging
import lxml.etree as LT


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


class ConditionalError(Error):
    """Raised when the input value to be derived from another field is left empty """
    pass


class NotImplementedError(Error):
    """Raised when code is not implemented for the type of subset type specified """
    pass


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


def aif_xml(files):
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
            output = os.path.splitext(file)[0]
            df = pd.read_excel(file, header=None)
            df = df.replace(np.nan, '', regex=True)
            df.columns = ['xmlTags', 'Id', 'XMLDescription', 'Input_1', 'Input_2', 'Input_3',
                          'Input_4', 'Input_5', 'Input_6', 'Input_7', 'Input_8', 'Input_9',
                          'Input_10', 'Input_11', 'Input_12']
            df.xmlTags = df.xmlTags.str.strip('<>')

#            print(df.head())

            headerRows = [str(i) for i in range(1, 4)]
            headerFileKeys = df[df.Id.isin(headerRows)][[
                'xmlTags', 'Input_1']].values.tolist()

            if headerFileKeys[0][1] and headerFileKeys[1][1] == "":
                raise EmptyValueError(
                    f"{headerFileKeys[0][0]} and {headerFileKeys[1][0]} fields cannot be empty!")

            generated_on = datetime.today().strftime('%Y-%m-%dT%H:%M:%S.0Z')
            root = ET.Element('AIFReportingInfo')
            root.set("xsi:noNamespaceSchemaLocation", "AIFMD_DATAIF_V1.2.xsd")
            root.set('CreationDateAndTime', generated_on)
            root.set(headerFileKeys[0][0], str(headerFileKeys[0][1]).strip())
            root.set(headerFileKeys[1][0], str(headerFileKeys[1][1]).strip())
            root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

            sectionRows = [str(i) for i in range(4, 10)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}

            AIFRecordInfo = ET.SubElement(root, 'AIFRecordInfo')
            for k, v in headerSectionKeys.items():
                if v != "":
                    if k not in ['ReportingPeriodStartDate', 'ReportingPeriodEndDate']:
                        k = ET.SubElement(AIFRecordInfo, k)
                        k.text = str(v).strip()
                    else:
                        k = ET.SubElement(AIFRecordInfo, k)
                        k.text = str(v.date()).strip()
                else:
                    raise EmptyValueError(f"{k} field cannot be empty!")

            sectionRows = [str(i) for i in range(10, 16)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}
#            print(headerSectionKeys)
#
            if headerSectionKeys['AIFReportingObligationChangeFrequencyCode'] != "":
                AIFReportingObligationChangeFrequencyCode = ET.SubElement(
                    AIFRecordInfo, 'AIFReportingObligationChangeFrequencyCode')
                AIFReportingObligationChangeFrequencyCode.text = str(
                    headerSectionKeys['AIFReportingObligationChangeFrequencyCode']).strip()

            if headerSectionKeys['AIFReportingObligationChangeContentsCode'] != "":
                AIFReportingObligationChangeContentsCode = ET.SubElement(
                    AIFRecordInfo, 'AIFReportingObligationChangeContentsCode')
                AIFReportingObligationChangeContentsCode.text = str(
                    headerSectionKeys['AIFReportingObligationChangeContentsCode']).strip()

            if headerSectionKeys['AIFReportingObligationChangeFrequencyCode'] or headerSectionKeys['AIFReportingObligationChangeContentsCode'] != "":
                if headerSectionKeys['AIFReportingObligationChangeQuarter'] != "":
                    AIFReportingObligationChangeQuarter = ET.SubElement(
                        AIFRecordInfo, 'AIFReportingObligationChangeQuarter')
                    AIFReportingObligationChangeQuarter.text = str(
                        headerSectionKeys['AIFReportingObligationChangeQuarter']).strip()
                else:
                    raise EmptyValueError(
                        "AIFReportingObligationChangeQuarter field cannot be empty!")

            if headerSectionKeys['LastReportingFlag'] != "":
                LastReportingFlag = ET.SubElement(
                    AIFRecordInfo, 'LastReportingFlag')
                LastReportingFlag.text = str(
                    headerSectionKeys['LastReportingFlag']).lower().strip()
            else:
                raise EmptyValueError(
                    "LastReportingFlag field cannot be empty!")

            if headerSectionKeys['QuestionNumber'] and headerSectionKeys['AssumptionDescription'] == "":
                if len(headerSectionKeys['AssumptionDescription']) > 300:
                    raise LengthValueRequiredError(
                        f'AssumptionDescription string required in this field should not be greater 300!')

                QuestionNumber = ET.SubElement(AIFRecordInfo, 'QuestionNumber')
                QuestionNumber.text = str(
                    headerSectionKeys['QuestionNumber']).strip()
                AssumptionDescription = ET.SubElement(
                    AIFRecordInfo, 'AssumptionDescription')
                AssumptionDescription.text = str(
                    headerSectionKeys['AssumptionDescription']).strip()

            sectionRows = [str(i) for i in range(16, 24)]
            headerSectionKeys = df[df.Id.isin(
                sectionRows)][['xmlTags', 'Input_1']].values.tolist()
            headerSectionKeys = {i[0]: i[1] for i in headerSectionKeys}

            for k, v in headerSectionKeys.items():
                if v != "":
                    k = ET.SubElement(AIFRecordInfo, k)
                    k.text = str(v).strip()
                else:
                    raise EmptyValueError(f"{k} field cannot be empty!")

            identifierRows = [str(i) for i in range(24, 33)]
            AIFIdentifiers = df[df.Id.isin(identifierRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            AIFIdentifiers = {i[0]: i[1].strip() for i in AIFIdentifiers}

            AIFCompleteDescription = ET.SubElement(
                AIFRecordInfo, 'AIFCompleteDescription')
            AIFPrincipalInfo = ET.SubElement(
                AIFCompleteDescription, 'AIFPrincipalInfo')

            if AIFIdentifiers:
                if (AIFIdentifiers['ReportingMemberState'] != "" and AIFIdentifiers['AIFNationalCode'] == "") or (AIFIdentifiers['AIFNationalCode'] != "" and AIFIdentifiers['ReportingMemberState'] == ""):
                    raise ConditionalError(
                        " Value required for ReportingMemberState if AIFNationalCode is filled and vice versa")

                counter = 0
                for k, v in AIFIdentifiers.items():
                    if v != "" and counter == 0:
                        AIFIdentification = ET.SubElement(
                            AIFPrincipalInfo, 'AIFIdentification')
                        k = ET.SubElement(AIFIdentification, k)
                        k.text = str(v).strip()
                    elif v != "" and counter != 0:
                        k = ET.SubElement(AIFIdentification, k)
                        k.text = str(v).strip()
                    counter += 1

            shareClassRows = [str(i) for i in range(33, 41)]
            shareClass = df[df.Id.isin(shareClassRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            shareClass = {i[0]: i[1] for i in shareClass}

            ShareClassFlag = ET.SubElement(AIFPrincipalInfo, 'ShareClassFlag')
#
            if str(shareClass['ShareClassFlag']).lower().strip() == 'false':
                ShareClassFlag.text = str(shareClass['ShareClassFlag']).strip()
            else:
                if shareClass['ShareClassName'] == "" and str(shareClass['ShareClassFlag']).lower().strip() == 'true':
                    raise EmptyValueError("Share class name field is required")

                ShareClassIdentification = ET.SubElement(
                    AIFPrincipalInfo, 'ShareClassIdentification')
                for k, v in shareClass.items():
                    ShareClassIdentifier = ET.SubElement(
                        ShareClassIdentification, 'ShareClassIdentifier')
                    if v != "":
                        k = ET.SubElement(ShareClassIdentifier, k)
                        k.text = str(v).strip()
#
            masterFeederRows = [str(i) for i in range(41, 45)]
            masterFeeder = df[df.Id.isin(masterFeederRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            masterFeeder = {i[0]: i[1] for i in masterFeeder}

            AIFDescription = ET.SubElement(AIFPrincipalInfo, 'AIFDescription')
            AIFMasterFeederStatus = ET.SubElement(
                AIFDescription, 'AIFMasterFeederStatus')
            AIFMasterFeederStatus.text = str(
                masterFeeder['AIFMasterFeederStatus']).upper().strip()

            if (masterFeeder['AIFMasterFeederStatus']).upper().strip() == "FEEDER":
                if masterFeeder['AIFName'] == "":
                    raise EmptyValueError('Value required for AIFName field')

                MasterAIFsIdentification = ET.SubElement(
                    AIFDescription, 'MasterAIFsIdentification')
                MasterAIFIdentification = ET.SubElement(
                    MasterAIFsIdentification, 'MasterAIFIdentification')

                AIFName = ET.SubElement(MasterAIFIdentification, 'AIFName')
                AIFName.text = str(masterFeeder['AIFName']).strip()

                if masterFeeder['ReportingMemberState'] != "" and masterFeeder['AIFNationalCode'] == "":
                    raise EmptyValueError(
                        "Value is required for AIFNationalCode field")

                AIFIdentifierNCA = ET.SubElement(
                    MasterAIFIdentification, 'AIFIdentifierNCA')

                if masterFeeder['ReportingMemberState'] != "":
                    ReportingMemberState = ET.SubElement(
                        AIFIdentifierNCA, 'ReportingMemberState')
                    ReportingMemberState.text = str(
                        masterFeeder['ReportingMemberState']).strip()
                if masterFeeder['AIFNationalCode'] != "":
                    AIFNationalCode = ET.SubElement(
                        AIFIdentifierNCA, 'AIFNationalCode')
                    AIFNationalCode.text = str(
                        masterFeeder['AIFNationalCode']).strip()

            primeBrokersRows = [str(i) for i in range(45, 48)]
            primeBrokers = df[df.Id.isin(primeBrokersRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            primeBrokers = {i[0]: i[1] for i in primeBrokers}

            counter = 0

            if primeBrokers:
                for k, v in AIFIdentifiers.items():
                    if v != "" and counter == 0:
                        PrimeBrokers = ET.SubElement(
                            AIFDescription, 'PrimeBrokers')
                        PrimeBrokerIdentification = ET.SubElement(
                            PrimeBrokers, 'PrimeBrokerIdentification')
                        k = ET.SubElement(
                            PrimeBrokerIdentification, k)
                        k.text = str(v).strip()
                    elif v != "" and counter != 0:
                        k = ET.SubElement(PrimeBrokerIdentification, k)
                        k.text = str(v).strip()
                    counter += 1

            valuesRows = (str(i) for i in range(48, 54))
            principalValues = df[df.Id.isin(
                valuesRows)][['xmlTags', 'Input_1']].values.tolist()
            principalValues = {i[0]: i[1] for i in principalValues}

            if principalValues['BaseCurrency'] and principalValues['AIFNetAssetValue'] and principalValues['AUMAmountInBaseCurrency'] == "":
                raise EmptyValueError(
                    ' AUMAmountInBaseCurrency, BaseCurrency & AIFNetAssetValue cannot be empty')

            AIFBaseCurrencyDescription = ET.SubElement(
                AIFDescription, 'AIFBaseCurrencyDescription')

            BaseCurrency = ET.SubElement(
                AIFBaseCurrencyDescription, 'BaseCurrency')
            BaseCurrency.text = str(
                principalValues['BaseCurrency']).upper().strip()
            AUMAmountInBaseCurrency = ET.SubElement(
                AIFBaseCurrencyDescription, 'AUMAmountInBaseCurrency')
            AUMAmountInBaseCurrency.text = str(
                principalValues['AUMAmountInBaseCurrency']).strip()

            if (principalValues['BaseCurrency']).upper().strip() != 'EUR':

                if principalValues['FXEURRate'] and principalValues['FXEURReferenceRateType'] != "":
                    FXEURReferenceRateType = ET.SubElement(
                        AIFBaseCurrencyDescription, 'FXEURReferenceRateType')
                    FXEURReferenceRateType.text = str(
                        principalValues['FXEURReferenceRateType']).upper().strip()
                    FXEURRate = ET.SubElement(
                        AIFBaseCurrencyDescription, 'FXEURRate')
                    FXEURRate.text = str(principalValues['FXEURRate']).strip()
                else:
                    raise EmptyValueError(
                        ' FXEURRate & FXEURReferenceRateType cannot be empty')

            if (principalValues['FXEURReferenceRateType']).upper().strip() == "OTH":
                if principalValues['FXEUROtherReferenceRateDescription'] != "":
                    FXEUROtherReferenceRateDescription = ET.SubElement(
                        AIFBaseCurrencyDescription, 'FXEUROtherReferenceRateDescription')
                    FXEUROtherReferenceRateDescription.text = str(
                        principalValues['FXEUROtherReferenceRateDescription']).strip()
                else:
                    raise EmptyValueError(
                        'FXEUROthReferenceRateDescription cannot be empty')

            AIFNetAssetValue = ET.SubElement(
                AIFDescription, 'AIFNetAssetValue')
            AIFNetAssetValue.text = str(
                principalValues['AIFNetAssetValue']).strip()
#
            jurisdictionRows = (str(i) for i in range(54, 58))
            jurisdictionValues = df[df.Id.isin(jurisdictionRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            jurisdictionValues = {i[0]: i[1] for i in jurisdictionValues}

            if jurisdictionValues['PredominantAIFType'] == "":
                raise EmptyValueError(
                    "PredominantAIFType field cannot be empty")

            if jurisdictionValues:
                for k, v in jurisdictionValues.items():
                    if v != "":
                        k = ET.SubElement(AIFDescription, k)
                        k.text = str(v).strip()
#
            investmentRows = [str(i) for i in range(58, 61)]
            investmentValues = df[df.Id.isin(investmentRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            investmentValues = {i[0]: i[1] for i in investmentValues}

            if jurisdictionValues['PredominantAIFType'] == "HFND":
                HedgeFundInvestmentStrategies = ET.SubElement(
                    AIFDescription, 'HedgeFundInvestmentStrategies')
                HedgeFundInvestmentStrategy = ET.SubElement(
                    HedgeFundInvestmentStrategies, 'HedgeFundInvestmentStrategy')
                for k, v in investmentValues.items():
                    k = ET.SubElement(HedgeFundInvestmentStrategy, k)
                    k.text = str(v).strip()
#
            elif jurisdictionValues['PredominantAIFType'] == "PEQF":
                PrivateEquityFundInvestmentStrategies = ET.SubElement(
                    AIFDescription, 'PrivateEquityFundInvestmentStrategies')
                PrivateEquityFundInvestmentStrategy = ET.SubElement(
                    PrivateEquityFundInvestmentStrategies, 'PrivateEquityFundInvestmentStrategy')
                for k, v in investmentValues.items():
                    if v != "":
                        k = ET.SubElement(
                            PrivateEquityFundInvestmentStrategy, k)
                        k.text = str(v).strip()
                    else:
                        raise EmptyValueError(
                            f"{k} field cannot be empty")
            else:
                raise NotImplementedError(
                    f"Script has not been implemented for this {investmentValues['PrivateEquityFundStrategyType']} yet")

            hFTTransactionNumber = [str(i) for i in range(62, 64)]
            hFTTransactionNumber = df[df.Id.isin(hFTTransactionNumber)][[
                'xmlTags', 'Input_1']].values.tolist()

            for FTTransactionNumber in hFTTransactionNumber:
                if FTTransactionNumber[1] != '':
                    FTTransactionNumber[0] = ET.SubElement(
                        AIFDescription, FTTransactionNumber[0])
                    FTTransactionNumber[0].text = str(
                        FTTransactionNumber[1]).strip()
#
            principalExRows = ['m' + str(i) for i in range(1, 6)]
            principalExValues = df[df.xmlTags.isin(principalExRows)][['Id', 'XMLDescription', 'Input_1', 'Input_2', 'Input_3',
                                                                      'Input_4', 'Input_5', 'Input_6', 'Input_7', 'Input_8', 'Input_9', 'Input_10', 'Input_11', 'Input_12']].values.tolist()
            MainInstrumentsTraded = ET.SubElement(
                AIFPrincipalInfo, 'MainInstrumentsTraded')

            for principalExValue in principalExValues:
                MainInstrumentTraded = ET.SubElement(
                    MainInstrumentsTraded, 'MainInstrumentTraded')
                Ranking = ET.SubElement(MainInstrumentTraded, 'Ranking')
                Ranking.text = str(principalExValue[0]).strip()
                if principalExValues[1] == "":
                    raise EmptyValueError("SubAssetType field cannot empty")
                SubAssetType = ET.SubElement(
                    MainInstrumentTraded, 'SubAssetType')
                SubAssetType.text = str(principalExValue[1]).strip()

                if principalExValue[1] != 'NTA_NTA_NOTA':
                    if principalExValue[2] == "":
                        raise EmptyValueError(
                            "InstrumentCodeType field cannot empty")
                    InstrumentCodeType = ET.SubElement(
                        MainInstrumentTraded, 'InstrumentCodeType')
                    InstrumentCodeType.text = str(
                        principalExValue[2]).strip()

                    if principalExValue[3] == "":
                        raise EmptyValueError(
                            "InstrumentName field cannot empty")
                    InstrumentName = ET.SubElement(
                        MainInstrumentTraded, 'InstrumentName')
                    InstrumentName.text = str(principalExValue[3]).strip()

                if principalExValue[2] == 'ISIN':
                    if principalExValue[4] == "":
                        raise EmptyValueError(
                            "ISINInstrumentIdentification field cannot empty")
                    ISINInstrumentIdentification = ET.SubElement(
                        MainInstrumentTraded, 'ISINInstrumentIdentification')
                    ISINInstrumentIdentification.text = str(
                        principalExValue[4]).strip()

                if principalExValue[2] == 'AII':
                    AIIInstrumentIdentification = AIIExchangeCode = ET.SubElement(
                        MainInstrumentTraded, 'AIIInstrumentIdentification')
                    print(principalExValue[9])
                    if principalExValue[5] == "":
                        raise EmptyValueError(
                            "AIIExchangeCode field cannot empty")
#
                    AIIExchangeCode = ET.SubElement(
                        AIIInstrumentIdentification, 'AIIExchangeCode')
                    AIIExchangeCode.text = str(
                        principalExValue[5]).strip()
                    if principalExValue[6] == "":
                        raise EmptyValueError(
                            "AIIDerivativeType field cannot empty")
                    AIIDerivativeType = ET.SubElement(
                        AIIInstrumentIdentification, 'AIIDerivativeType')
                    AIIDerivativeType.text = str(
                        principalExValue[6]).strip()
                    if principalExValue[7] == "":
                        raise EmptyValueError(
                            "AIIPutCallIdentifier field cannot empty")
                    AIIPutCallIdentifier = ET.SubElement(
                        AIIInstrumentIdentification, 'AIIPutCallIdentifier')
                    AIIPutCallIdentifier.text = str(
                        principalExValue[7]).strip()
                    if principalExValue[8] == "":
                        raise EmptyValueError(
                            "AIIExpiryDate field cannot empty")
                    AIIExpiryDate = ET.SubElement(
                        AIIInstrumentIdentification, 'AIIExpiryDate')
                    AIIExpiryDate.text = str(
                        principalExValue[8]).strip()
                    if principalExValue[9] == "":
                        raise EmptyValueError(
                            "AIIStrikePrice field cannot empty")
                    AIIStrikePrice = ET.SubElement(
                        AIIInstrumentIdentification, 'AIIStrikePrice')
                    AIIStrikePrice.text = str(
                        principalExValue[9]).strip()

                if principalExValue[1] != 'NTA_NTA_NOTA':
                    PositionValue = ET.SubElement(
                        MainInstrumentTraded, 'PositionValue')
                    PositionValue.text = str(principalExValue[-2]).strip()
                    PositionType = ET.SubElement(
                        MainInstrumentTraded, 'PositionType')
                    PositionType.text = str(
                        principalExValue[-3]).upper().strip()

                if str(principalExValue[-3]).upper().strip() == 'S':
                    ShortPositionHedgingRate = ET.SubElement(
                        MainInstrumentTraded, 'ShortPositionHedgingRate')
                    ShortPositionHedgingRate.text = str(
                        principalExValue[-1]).strip()

            NAVGeographicalFocusRows = [str(i) for i in range(78, 86)]
            navGeographicalFocus = df[df.Id.isin(NAVGeographicalFocusRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            navGeographicalFocus = {i[0]: i[1] for i in navGeographicalFocus}
            NAVGeographicalFocus = ET.SubElement(
                AIFPrincipalInfo, 'NAVGeographicalFocus')

            for k, v in navGeographicalFocus.items():
                if v != "":
                    k = ET.SubElement(NAVGeographicalFocus, k)
                    k.text = str(v).strip()
                else:
                    raise EmptyValueError(f"{k} cannot be empty")
#
            AUMGeographicalFocusRows = [str(i) for i in range(86, 94)]
            aumGeographicalFocus = df[df.Id.isin(AUMGeographicalFocusRows)][[
                'xmlTags', 'Input_1']].values.tolist()
            aumGeographicalFocus = {i[0]: i[1] for i in aumGeographicalFocus}
            if aumGeographicalFocus:
                AUMGeographicalFocus = ET.SubElement(
                    AIFPrincipalInfo, 'AUMGeographicalFocus')
                for k, v in aumGeographicalFocus.items():
                    k = ET.SubElement(AUMGeographicalFocus, k)
                    k.text = str(v).strip()
#
            principalEx2Values = ['p' + str(i) for i in range(1, 11)]
            principalEx2Values = df[df.xmlTags.isin(principalEx2Values)][['Id', 'XMLDescription', 'Input_1', 'Input_2', 'Input_3',
                                                                          'Input_4', 'Input_5', 'Input_6', 'Input_7']].values.tolist()
            PrincipalExposures = ET.SubElement(
                AIFPrincipalInfo, 'PrincipalExposures')

            for principalEx2Value in principalEx2Values:
                PrincipalExposure = ET.SubElement(
                    PrincipalExposures, 'PrincipalExposure')
                Ranking = ET.SubElement(PrincipalExposure, 'Ranking')
                Ranking.text = str(principalEx2Value[0]).strip()
                AssetMacroType = ET.SubElement(
                    PrincipalExposure, 'AssetMacroType')
                AssetMacroType.text = str(principalEx2Value[1]).strip()

                if principalEx2Value[1] != 'NTA':
                    SubAssetType = ET.SubElement(
                        PrincipalExposure, 'SubAssetType')
                    SubAssetType.text = str(principalEx2Value[2]).strip()
                    PositionType = ET.SubElement(
                        PrincipalExposure, 'PositionType')
                    PositionType.text = str(principalEx2Value[3]).strip()
                    AggregatedValueAmount = ET.SubElement(
                        PrincipalExposure, 'AggregatedValueAmount')
                    AggregatedValueAmount.text = str(
                        principalEx2Value[4]).strip()
                    AggregatedValueRate = ET.SubElement(
                        PrincipalExposure, 'AggregatedValueRate')
                    AggregatedValueRate.text = str(
                        principalEx2Value[5]).strip()

                    if principalEx2Value[6] != '':
                        CounterpartyIdentification = ET.SubElement(
                            PrincipalExposure, 'CounterpartyIdentification')
                        EntityName = ET.SubElement(
                            CounterpartyIdentification, 'EntityName')
                        EntityName.text = str(principalEx2Value[6]).strip()

                    if principalEx2Value[8] != '':
                        EntityIdentificationBIC = ET.SubElement(
                            CounterpartyIdentification, 'EntityIdentificationBIC')
                        EntityIdentificationBIC.text = str(
                            principalEx2Value[8]).strip()

                    if principalEx2Value[7] != '':
                        EntityIdentificationLEI = ET.SubElement(
                            CounterpartyIdentification, 'EntityIdentificationLEI')
                        EntityIdentificationLEI.text = str(
                            principalEx2Value[7]).strip()
#
            portfolioConcentration = ['q' + str(i) for i in range(1, 6)]
            portfolioConcentration = df[df.xmlTags.isin(portfolioConcentration)][['Id', 'XMLDescription', 'Input_1', 'Input_2', 'Input_3',
                                                                                  'Input_4', 'Input_5', 'Input_6', 'Input_7', 'Input_8']].values.tolist()
            MostImportantConcentration = ET.SubElement(
                AIFPrincipalInfo, 'MostImportantConcentration')
            PortfolioConcentrations = ET.SubElement(
                MostImportantConcentration, 'PortfolioConcentrations')

            for value in portfolioConcentration:
                PortfolioConcentration = ET.SubElement(
                    PortfolioConcentrations, 'PortfolioConcentration')
                Ranking = ET.SubElement(PortfolioConcentration, 'Ranking')
                Ranking.text = str(value[0]).strip()
                AssetType = ET.SubElement(
                    PortfolioConcentration, 'AssetType')
                AssetType.text = str(value[1]).strip()

                if value[1] != 'NTA_NTA':
                    PositionType = ET.SubElement(
                        PortfolioConcentration, 'PositionType')
                    PositionType.text = str(value[2]).strip()
                    MarketIdentification = ET.SubElement(
                        PortfolioConcentration, 'MarketIdentification')
                    MarketCodeType = ET.SubElement(
                        MarketIdentification, 'MarketCodeType')
                    MarketCodeType.text = str(value[3]).strip()

                if value[3] == "MIC":
                    MarketCode = ET.SubElement(
                        MarketIdentification, 'MarketCode')
                    MarketCode.text = str(value[4]).strip()

                AggregatedValueAmount = ET.SubElement(
                    PortfolioConcentration, 'AggregatedValueAmount')
                AggregatedValueAmount.text = str(value[5]).strip()
                AggregatedValueRate = ET.SubElement(
                    PortfolioConcentration, 'AggregatedValueRate')
                AggregatedValueRate.text = str(value[6]).strip()

                if value[3] == 'OTC' and value[7] != "":
                    CounterpartyIdentification = ET.SubElement(
                        PortfolioConcentration, 'CounterpartyIdentification')
                    EntityName = ET.SubElement(
                        CounterpartyIdentification, 'EntityName')
                    EntityName.text = str(value[7]).strip()

                if value[9] != '':
                    EntityIdentificationBIC = ET.SubElement(
                        CounterpartyIdentification, 'EntityIdentificationBIC')
                    EntityIdentificationBIC.text = str(
                        value[9]).strip()
                if value[8] != '':
                    EntityIdentificationLEI = ET.SubElement(
                        CounterpartyIdentification, 'EntityIdentificationLEI')
                    EntityIdentificationLEI.text = str(
                        value[8]).strip()

            typicalPositionSize = df[df.Id == '113']['Input_1'].values.tolist()[
                0]
            if jurisdictionValues['PredominantAIFType'] == "PEQF":
                TypicalPositionSize = ET.SubElement(
                    MostImportantConcentration, 'TypicalPositionSize')
                TypicalPositionSize.text = str(typicalPositionSize).strip()
#
            markerts = ['r' + str(i) for i in range(1, 4)]
            markerts = df[df.xmlTags.isin(
                markerts)][['Id', 'XMLDescription', 'Input_1', 'Input_2']].values.tolist()
            AIFPrincipalMarkets = ET.SubElement(
                MostImportantConcentration, 'AIFPrincipalMarkets')

            for market in markerts:
                AIFPrincipalMarket = ET.SubElement(
                    AIFPrincipalMarkets, 'AIFPrincipalMarket')
                Ranking = ET.SubElement(AIFPrincipalMarket, 'Ranking')
                Ranking.text = str(market[0]).strip()

                if market[1] == "":
                    raise EmptyValueError(
                        "MarketIdentification cannot be empty")
                MarketIdentification = ET.SubElement(
                    AIFPrincipalMarket, 'MarketIdentification')
                MarketCodeType = ET.SubElement(
                    MarketIdentification, 'MarketCodeType')
                MarketCodeType.text = str(market[1]).upper().strip()

                if str(market[1]).upper().strip() == 'MIC' and market[2] == "":
                    raise EmptyValueError("MarketCode cannot be empty")

                if str(market[1]).upper().strip() == 'MIC':
                    MarketCode = ET.SubElement(
                        MarketIdentification, 'MarketCode')
                    MarketCode.text = str(market[2]).strip()

                if str(market[1]).upper().strip() != 'NOT' and market[3] == "":
                    raise EmptyValueError("MarketCode cannot be empty")

                if str(market[1]).upper().strip() != 'NOT':
                    AggregatedValueAmount = ET.SubElement(
                        AIFPrincipalMarket, 'AggregatedValueAmount')
                    AggregatedValueAmount.text = str(market[3]).strip()

            investorConcentration = [str(i) for i in range(118, 121)]
            investorConcentration = df[df.Id.isin(investorConcentration)][[
                'xmlTags', 'Input_1']].values.tolist()
            investorConcentration = {i[0]: i[1] for i in investorConcentration}

            InvestorConcentration = ET.SubElement(
                MostImportantConcentration, 'InvestorConcentration')

            for k, v in investorConcentration.items():
                if v != "":
                    k = ET.SubElement(InvestorConcentration, k)
                    k.text = str(v).strip()
                else:
                    raise EmptyValueError(f"{k} cannot be empty")

        tree = ET.ElementTree(root)
        tree.write(output + '.xml', encoding="UTF-8", xml_declaration=True)

    except Exception as e:
        print(e)


def validate_XML_AIF(files):
    for file in files:
        xml_file = LT.parse(file)
        xml_validator = LT.XMLSchema(file="AIFMD_DATAIF_V1.2.xsd")
        is_valid = xml_validator.validate(xml_file)
        if is_valid:
            logger.debug(f'{file} has successfully been validated!')
        else:
            logger.error(f'Validation for {file} was unsuccessful!')


if __name__ == '__main__':
    path = 'Enter file path'
    os.chdir(path)
    files = []
    types = ('*.xls', '*.xlsx')
    for ext in types:
        files.extend(glob.glob(ext))
    aif_xml(files)
    xmls = glob.glob("*.xml")
    validate_XML_AIF(xmls)
