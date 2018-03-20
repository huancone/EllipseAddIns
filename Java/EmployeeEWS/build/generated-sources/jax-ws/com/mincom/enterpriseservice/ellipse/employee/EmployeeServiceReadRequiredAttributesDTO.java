
package com.mincom.enterpriseservice.ellipse.employee;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlType;
import com.mincom.enterpriseservice.ellipse.AbstractRequiredAttributesDTO;


/**
 * <p>Java class for EmployeeServiceReadRequiredAttributesDTO complex type.
 * 
 * <p>The following schema fragment specifies the expected content contained within this class.
 * 
 * <pre>
 * &lt;complexType name="EmployeeServiceReadRequiredAttributesDTO">
 *   &lt;complexContent>
 *     &lt;extension base="{http://ellipse.enterpriseservice.mincom.com}AbstractRequiredAttributesDTO">
 *       &lt;sequence>
 *         &lt;element name="returnActualFTEPercent" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAuthorityPercent" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAwardCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnAwardCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnBarcode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnBirthDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnBonaFideTermination" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCandidateId" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCitizenIndicator" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCitizenIndicatorDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCompetencyDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCompetencyLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnContractHours" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnContractMinutes" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCopyResAddrPostal" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCoreEmployeeInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCountryOfBirth" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnCountryOfBirthDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDataRefRequired" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDataReferenceNo" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDeathReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDependants" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDisabledInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnDuplicateNameInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmailAddress" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployee" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeClass" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeClassDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeFormattedName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEmployeeTypeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEntitleDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEntitleId" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEssUserInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEthnicity" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnEthnicityDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnExcludeTalentExtract" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnFirstName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnFixedAssetsDistrict" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnFixedAssetsDistrictDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnGender" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnGenderDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnGlobalProfile" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHealthPlan" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHireDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHomeFacsimileNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHomeMobilePhoneNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnHomeTelephoneNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnJobClassLevel" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnJobClassLevelDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLanguageCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLanguageCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLastName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLeaveForecastDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLegalRepIndicator" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnLegalRepName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMaritalStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMaritalStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMessagePreference" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnMessagePreferenceDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdCode1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdCode2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdCode3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdCode4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdCode5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType1Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType2Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType3Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType4Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalIdType5Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationality" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNationalityDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnNotifyEDIMsgRecieved" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportExpiryDate1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportExpiryDate2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportExpiryDate3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportIssuedBy1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportIssuedBy2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportIssuedBy3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportNumber1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportNumber2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPassportNumber3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPaygroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPaygroupDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPayrollEmployeeInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersEmpStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersEmpStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonTypeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonalEmail" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass10" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass10Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass10Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass1Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass1Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass2Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass2Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass3Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass3Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass4Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass4Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass5Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass5Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass6" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass6Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass6Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass7" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass7Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass7Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass8" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass8Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass8Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass9" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass9Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelClass9Name" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelEmployeeInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelGroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelGroupDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPersonnelStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhotoPathname" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocation" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPhysicalLocationDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPosition" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionClassUDef1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionClassUDef2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionClassUDefName1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionClassUDefName2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPositionStartDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalAddressLine1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalAddressLine2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalAddressLine3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalCountry" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalCountryDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalState" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalStateDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPostalZipCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreferredName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreviousEmployeeId" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreviousFirstName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreviousLastName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreviousSecondName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPreviousThirdName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrimRepCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterCode5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterDesc5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterName1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterName2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterName3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterName4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnPrinterName5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnProfessionalServiceDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnRehireCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnRehireCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnReinstatementDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialAddressEffDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialAddressLine1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialAddressLine2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialAddressLine3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialCountry" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialCountryDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialState" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialStateDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResidentialZipCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceClass" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnResourceEffectiveDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnRetirementDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSearchStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSeasonalInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSeasonalIndDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSecondName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnServiceDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReason1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReason2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReason3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReason4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReasonDesc1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReasonDesc2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReasonDesc3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSigDateReasonDesc4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSignificantDate1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSignificantDate2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSignificantDate3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSignificantDate4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSiteCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSiteCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportCode1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportCode2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportCode3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportType1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportType2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSkillsPassportType3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSmokerInd" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialInsuranceNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialSecurityNoType" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialSecurityNoTypeDescription" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSocialSecurityNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnStaffCategory" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnStaffCategoryDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnStdTextRef" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnStdTextRefExists" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSupplier" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSupplierName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnSuspensionDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTerminationDate" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTerminationReason" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTerminationReasonDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnThirdName" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTitle" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnTitleDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUnionCode" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUnionCodeDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUserDefContact1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUserDefContact2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUserDefContact3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUserDefContact4" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnUserDefContact5" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVeteranStatus" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVeteranStatusDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode1Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode2Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaCode3Desc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaEffectiveDate1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaEffectiveDate2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaEffectiveDate3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaExpiryDate1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaExpiryDate2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaExpiryDate3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaIssuedBy1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaIssuedBy2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaIssuedBy3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaNumber1" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaNumber2" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnVisaNumber3" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkFacsimileNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroup" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupCrew" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupCrewDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkGroupDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkLocation" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkLocationDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkMobilePhoneNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkOrderPrefix" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkOrderPrefixDesc" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkTelephoneExtension" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *         &lt;element name="returnWorkTelephoneNumber" type="{http://www.w3.org/2001/XMLSchema}boolean" minOccurs="0"/>
 *       &lt;/sequence>
 *     &lt;/extension>
 *   &lt;/complexContent>
 * &lt;/complexType>
 * </pre>
 * 
 * 
 */
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "EmployeeServiceReadRequiredAttributesDTO", propOrder = {
    "returnActualFTEPercent",
    "returnAuthorityPercent",
    "returnAwardCode",
    "returnAwardCodeDesc",
    "returnBarcode",
    "returnBirthDate",
    "returnBonaFideTermination",
    "returnCandidateId",
    "returnCitizenIndicator",
    "returnCitizenIndicatorDesc",
    "returnCompetencyDesc",
    "returnCompetencyLevel",
    "returnContractHours",
    "returnContractMinutes",
    "returnCopyResAddrPostal",
    "returnCoreEmployeeInd",
    "returnCountryOfBirth",
    "returnCountryOfBirthDesc",
    "returnDataRefRequired",
    "returnDataReferenceNo",
    "returnDeathDate",
    "returnDeathReason",
    "returnDeathReasonDesc",
    "returnDependants",
    "returnDisabledInd",
    "returnDuplicateNameInd",
    "returnEmailAddress",
    "returnEmployee",
    "returnEmployeeClass",
    "returnEmployeeClassDesc",
    "returnEmployeeFormattedName",
    "returnEmployeeType",
    "returnEmployeeTypeDesc",
    "returnEntitleDesc",
    "returnEntitleId",
    "returnEssUserInd",
    "returnEthnicity",
    "returnEthnicityDesc",
    "returnExcludeTalentExtract",
    "returnFirstName",
    "returnFixedAssetsDistrict",
    "returnFixedAssetsDistrictDesc",
    "returnGender",
    "returnGenderDesc",
    "returnGlobalProfile",
    "returnHealthPlan",
    "returnHireDate",
    "returnHomeFacsimileNumber",
    "returnHomeMobilePhoneNumber",
    "returnHomeTelephoneNumber",
    "returnJobClassLevel",
    "returnJobClassLevelDesc",
    "returnLanguageCode",
    "returnLanguageCodeDesc",
    "returnLastName",
    "returnLeaveForecastDate",
    "returnLegalRepIndicator",
    "returnLegalRepName",
    "returnMaritalStatus",
    "returnMaritalStatusDesc",
    "returnMessagePreference",
    "returnMessagePreferenceDesc",
    "returnNationalIdCode1",
    "returnNationalIdCode2",
    "returnNationalIdCode3",
    "returnNationalIdCode4",
    "returnNationalIdCode5",
    "returnNationalIdType1",
    "returnNationalIdType1Desc",
    "returnNationalIdType2",
    "returnNationalIdType2Desc",
    "returnNationalIdType3",
    "returnNationalIdType3Desc",
    "returnNationalIdType4",
    "returnNationalIdType4Desc",
    "returnNationalIdType5",
    "returnNationalIdType5Desc",
    "returnNationality",
    "returnNationalityDesc",
    "returnNotifyEDIMsgRecieved",
    "returnPassportExpiryDate1",
    "returnPassportExpiryDate2",
    "returnPassportExpiryDate3",
    "returnPassportIssuedBy1",
    "returnPassportIssuedBy2",
    "returnPassportIssuedBy3",
    "returnPassportNumber1",
    "returnPassportNumber2",
    "returnPassportNumber3",
    "returnPaygroup",
    "returnPaygroupDesc",
    "returnPayrollEmployeeInd",
    "returnPersEmpStatus",
    "returnPersEmpStatusDesc",
    "returnPersonType",
    "returnPersonTypeDesc",
    "returnPersonalEmail",
    "returnPersonnelClass1",
    "returnPersonnelClass10",
    "returnPersonnelClass10Desc",
    "returnPersonnelClass10Name",
    "returnPersonnelClass1Desc",
    "returnPersonnelClass1Name",
    "returnPersonnelClass2",
    "returnPersonnelClass2Desc",
    "returnPersonnelClass2Name",
    "returnPersonnelClass3",
    "returnPersonnelClass3Desc",
    "returnPersonnelClass3Name",
    "returnPersonnelClass4",
    "returnPersonnelClass4Desc",
    "returnPersonnelClass4Name",
    "returnPersonnelClass5",
    "returnPersonnelClass5Desc",
    "returnPersonnelClass5Name",
    "returnPersonnelClass6",
    "returnPersonnelClass6Desc",
    "returnPersonnelClass6Name",
    "returnPersonnelClass7",
    "returnPersonnelClass7Desc",
    "returnPersonnelClass7Name",
    "returnPersonnelClass8",
    "returnPersonnelClass8Desc",
    "returnPersonnelClass8Name",
    "returnPersonnelClass9",
    "returnPersonnelClass9Desc",
    "returnPersonnelClass9Name",
    "returnPersonnelEmployeeInd",
    "returnPersonnelGroup",
    "returnPersonnelGroupDesc",
    "returnPersonnelStatus",
    "returnPersonnelStatusDesc",
    "returnPhotoPathname",
    "returnPhysicalLocReason",
    "returnPhysicalLocReasonDesc",
    "returnPhysicalLocation",
    "returnPhysicalLocationDesc",
    "returnPosition",
    "returnPositionClassUDef1",
    "returnPositionClassUDef2",
    "returnPositionClassUDefName1",
    "returnPositionClassUDefName2",
    "returnPositionDesc",
    "returnPositionReason",
    "returnPositionReasonDesc",
    "returnPositionStartDate",
    "returnPostalAddressLine1",
    "returnPostalAddressLine2",
    "returnPostalAddressLine3",
    "returnPostalCountry",
    "returnPostalCountryDesc",
    "returnPostalState",
    "returnPostalStateDesc",
    "returnPostalZipCode",
    "returnPreferredName",
    "returnPreviousEmployeeId",
    "returnPreviousFirstName",
    "returnPreviousLastName",
    "returnPreviousSecondName",
    "returnPreviousThirdName",
    "returnPrimRepCode",
    "returnPrimRepCodeDesc",
    "returnPrinterCode1",
    "returnPrinterCode2",
    "returnPrinterCode3",
    "returnPrinterCode4",
    "returnPrinterCode5",
    "returnPrinterDesc1",
    "returnPrinterDesc2",
    "returnPrinterDesc3",
    "returnPrinterDesc4",
    "returnPrinterDesc5",
    "returnPrinterName1",
    "returnPrinterName2",
    "returnPrinterName3",
    "returnPrinterName4",
    "returnPrinterName5",
    "returnProfessionalServiceDate",
    "returnRehireCode",
    "returnRehireCodeDesc",
    "returnReinstatementDate",
    "returnResidentialAddressEffDate",
    "returnResidentialAddressLine1",
    "returnResidentialAddressLine2",
    "returnResidentialAddressLine3",
    "returnResidentialCountry",
    "returnResidentialCountryDesc",
    "returnResidentialState",
    "returnResidentialStateDesc",
    "returnResidentialZipCode",
    "returnResourceClass",
    "returnResourceCode",
    "returnResourceDesc",
    "returnResourceEffectiveDate",
    "returnRetirementDate",
    "returnSearchStatus",
    "returnSeasonalInd",
    "returnSeasonalIndDesc",
    "returnSecondName",
    "returnServiceDate",
    "returnSigDateReason1",
    "returnSigDateReason2",
    "returnSigDateReason3",
    "returnSigDateReason4",
    "returnSigDateReasonDesc1",
    "returnSigDateReasonDesc2",
    "returnSigDateReasonDesc3",
    "returnSigDateReasonDesc4",
    "returnSignificantDate1",
    "returnSignificantDate2",
    "returnSignificantDate3",
    "returnSignificantDate4",
    "returnSiteCode",
    "returnSiteCodeDesc",
    "returnSkillsPassportCode1",
    "returnSkillsPassportCode2",
    "returnSkillsPassportCode3",
    "returnSkillsPassportType1",
    "returnSkillsPassportType2",
    "returnSkillsPassportType3",
    "returnSmokerInd",
    "returnSocialInsuranceNumber",
    "returnSocialSecurityNoType",
    "returnSocialSecurityNoTypeDescription",
    "returnSocialSecurityNumber",
    "returnStaffCategory",
    "returnStaffCategoryDesc",
    "returnStdTextRef",
    "returnStdTextRefExists",
    "returnSupplier",
    "returnSupplierName",
    "returnSuspensionDate",
    "returnTerminationDate",
    "returnTerminationReason",
    "returnTerminationReasonDesc",
    "returnThirdName",
    "returnTitle",
    "returnTitleDesc",
    "returnUnionCode",
    "returnUnionCodeDesc",
    "returnUserDefContact1",
    "returnUserDefContact2",
    "returnUserDefContact3",
    "returnUserDefContact4",
    "returnUserDefContact5",
    "returnVeteranStatus",
    "returnVeteranStatusDesc",
    "returnVisaCode1",
    "returnVisaCode1Desc",
    "returnVisaCode2",
    "returnVisaCode2Desc",
    "returnVisaCode3",
    "returnVisaCode3Desc",
    "returnVisaEffectiveDate1",
    "returnVisaEffectiveDate2",
    "returnVisaEffectiveDate3",
    "returnVisaExpiryDate1",
    "returnVisaExpiryDate2",
    "returnVisaExpiryDate3",
    "returnVisaIssuedBy1",
    "returnVisaIssuedBy2",
    "returnVisaIssuedBy3",
    "returnVisaNumber1",
    "returnVisaNumber2",
    "returnVisaNumber3",
    "returnWorkFacsimileNumber",
    "returnWorkGroup",
    "returnWorkGroupCrew",
    "returnWorkGroupCrewDesc",
    "returnWorkGroupDesc",
    "returnWorkLocation",
    "returnWorkLocationDesc",
    "returnWorkMobilePhoneNumber",
    "returnWorkOrderPrefix",
    "returnWorkOrderPrefixDesc",
    "returnWorkTelephoneExtension",
    "returnWorkTelephoneNumber"
})
public class EmployeeServiceReadRequiredAttributesDTO
    extends AbstractRequiredAttributesDTO
{

    protected Boolean returnActualFTEPercent;
    protected Boolean returnAuthorityPercent;
    protected Boolean returnAwardCode;
    protected Boolean returnAwardCodeDesc;
    protected Boolean returnBarcode;
    protected Boolean returnBirthDate;
    protected Boolean returnBonaFideTermination;
    protected Boolean returnCandidateId;
    protected Boolean returnCitizenIndicator;
    protected Boolean returnCitizenIndicatorDesc;
    protected Boolean returnCompetencyDesc;
    protected Boolean returnCompetencyLevel;
    protected Boolean returnContractHours;
    protected Boolean returnContractMinutes;
    protected Boolean returnCopyResAddrPostal;
    protected Boolean returnCoreEmployeeInd;
    protected Boolean returnCountryOfBirth;
    protected Boolean returnCountryOfBirthDesc;
    protected Boolean returnDataRefRequired;
    protected Boolean returnDataReferenceNo;
    protected Boolean returnDeathDate;
    protected Boolean returnDeathReason;
    protected Boolean returnDeathReasonDesc;
    protected Boolean returnDependants;
    protected Boolean returnDisabledInd;
    protected Boolean returnDuplicateNameInd;
    protected Boolean returnEmailAddress;
    protected Boolean returnEmployee;
    protected Boolean returnEmployeeClass;
    protected Boolean returnEmployeeClassDesc;
    protected Boolean returnEmployeeFormattedName;
    protected Boolean returnEmployeeType;
    protected Boolean returnEmployeeTypeDesc;
    protected Boolean returnEntitleDesc;
    protected Boolean returnEntitleId;
    protected Boolean returnEssUserInd;
    protected Boolean returnEthnicity;
    protected Boolean returnEthnicityDesc;
    protected Boolean returnExcludeTalentExtract;
    protected Boolean returnFirstName;
    protected Boolean returnFixedAssetsDistrict;
    protected Boolean returnFixedAssetsDistrictDesc;
    protected Boolean returnGender;
    protected Boolean returnGenderDesc;
    protected Boolean returnGlobalProfile;
    protected Boolean returnHealthPlan;
    protected Boolean returnHireDate;
    protected Boolean returnHomeFacsimileNumber;
    protected Boolean returnHomeMobilePhoneNumber;
    protected Boolean returnHomeTelephoneNumber;
    protected Boolean returnJobClassLevel;
    protected Boolean returnJobClassLevelDesc;
    protected Boolean returnLanguageCode;
    protected Boolean returnLanguageCodeDesc;
    protected Boolean returnLastName;
    protected Boolean returnLeaveForecastDate;
    protected Boolean returnLegalRepIndicator;
    protected Boolean returnLegalRepName;
    protected Boolean returnMaritalStatus;
    protected Boolean returnMaritalStatusDesc;
    protected Boolean returnMessagePreference;
    protected Boolean returnMessagePreferenceDesc;
    protected Boolean returnNationalIdCode1;
    protected Boolean returnNationalIdCode2;
    protected Boolean returnNationalIdCode3;
    protected Boolean returnNationalIdCode4;
    protected Boolean returnNationalIdCode5;
    protected Boolean returnNationalIdType1;
    protected Boolean returnNationalIdType1Desc;
    protected Boolean returnNationalIdType2;
    protected Boolean returnNationalIdType2Desc;
    protected Boolean returnNationalIdType3;
    protected Boolean returnNationalIdType3Desc;
    protected Boolean returnNationalIdType4;
    protected Boolean returnNationalIdType4Desc;
    protected Boolean returnNationalIdType5;
    protected Boolean returnNationalIdType5Desc;
    protected Boolean returnNationality;
    protected Boolean returnNationalityDesc;
    protected Boolean returnNotifyEDIMsgRecieved;
    protected Boolean returnPassportExpiryDate1;
    protected Boolean returnPassportExpiryDate2;
    protected Boolean returnPassportExpiryDate3;
    protected Boolean returnPassportIssuedBy1;
    protected Boolean returnPassportIssuedBy2;
    protected Boolean returnPassportIssuedBy3;
    protected Boolean returnPassportNumber1;
    protected Boolean returnPassportNumber2;
    protected Boolean returnPassportNumber3;
    protected Boolean returnPaygroup;
    protected Boolean returnPaygroupDesc;
    protected Boolean returnPayrollEmployeeInd;
    protected Boolean returnPersEmpStatus;
    protected Boolean returnPersEmpStatusDesc;
    protected Boolean returnPersonType;
    protected Boolean returnPersonTypeDesc;
    protected Boolean returnPersonalEmail;
    protected Boolean returnPersonnelClass1;
    protected Boolean returnPersonnelClass10;
    protected Boolean returnPersonnelClass10Desc;
    protected Boolean returnPersonnelClass10Name;
    protected Boolean returnPersonnelClass1Desc;
    protected Boolean returnPersonnelClass1Name;
    protected Boolean returnPersonnelClass2;
    protected Boolean returnPersonnelClass2Desc;
    protected Boolean returnPersonnelClass2Name;
    protected Boolean returnPersonnelClass3;
    protected Boolean returnPersonnelClass3Desc;
    protected Boolean returnPersonnelClass3Name;
    protected Boolean returnPersonnelClass4;
    protected Boolean returnPersonnelClass4Desc;
    protected Boolean returnPersonnelClass4Name;
    protected Boolean returnPersonnelClass5;
    protected Boolean returnPersonnelClass5Desc;
    protected Boolean returnPersonnelClass5Name;
    protected Boolean returnPersonnelClass6;
    protected Boolean returnPersonnelClass6Desc;
    protected Boolean returnPersonnelClass6Name;
    protected Boolean returnPersonnelClass7;
    protected Boolean returnPersonnelClass7Desc;
    protected Boolean returnPersonnelClass7Name;
    protected Boolean returnPersonnelClass8;
    protected Boolean returnPersonnelClass8Desc;
    protected Boolean returnPersonnelClass8Name;
    protected Boolean returnPersonnelClass9;
    protected Boolean returnPersonnelClass9Desc;
    protected Boolean returnPersonnelClass9Name;
    protected Boolean returnPersonnelEmployeeInd;
    protected Boolean returnPersonnelGroup;
    protected Boolean returnPersonnelGroupDesc;
    protected Boolean returnPersonnelStatus;
    protected Boolean returnPersonnelStatusDesc;
    protected Boolean returnPhotoPathname;
    protected Boolean returnPhysicalLocReason;
    protected Boolean returnPhysicalLocReasonDesc;
    protected Boolean returnPhysicalLocation;
    protected Boolean returnPhysicalLocationDesc;
    protected Boolean returnPosition;
    protected Boolean returnPositionClassUDef1;
    protected Boolean returnPositionClassUDef2;
    protected Boolean returnPositionClassUDefName1;
    protected Boolean returnPositionClassUDefName2;
    protected Boolean returnPositionDesc;
    protected Boolean returnPositionReason;
    protected Boolean returnPositionReasonDesc;
    protected Boolean returnPositionStartDate;
    protected Boolean returnPostalAddressLine1;
    protected Boolean returnPostalAddressLine2;
    protected Boolean returnPostalAddressLine3;
    protected Boolean returnPostalCountry;
    protected Boolean returnPostalCountryDesc;
    protected Boolean returnPostalState;
    protected Boolean returnPostalStateDesc;
    protected Boolean returnPostalZipCode;
    protected Boolean returnPreferredName;
    protected Boolean returnPreviousEmployeeId;
    protected Boolean returnPreviousFirstName;
    protected Boolean returnPreviousLastName;
    protected Boolean returnPreviousSecondName;
    protected Boolean returnPreviousThirdName;
    protected Boolean returnPrimRepCode;
    protected Boolean returnPrimRepCodeDesc;
    protected Boolean returnPrinterCode1;
    protected Boolean returnPrinterCode2;
    protected Boolean returnPrinterCode3;
    protected Boolean returnPrinterCode4;
    protected Boolean returnPrinterCode5;
    protected Boolean returnPrinterDesc1;
    protected Boolean returnPrinterDesc2;
    protected Boolean returnPrinterDesc3;
    protected Boolean returnPrinterDesc4;
    protected Boolean returnPrinterDesc5;
    protected Boolean returnPrinterName1;
    protected Boolean returnPrinterName2;
    protected Boolean returnPrinterName3;
    protected Boolean returnPrinterName4;
    protected Boolean returnPrinterName5;
    protected Boolean returnProfessionalServiceDate;
    protected Boolean returnRehireCode;
    protected Boolean returnRehireCodeDesc;
    protected Boolean returnReinstatementDate;
    protected Boolean returnResidentialAddressEffDate;
    protected Boolean returnResidentialAddressLine1;
    protected Boolean returnResidentialAddressLine2;
    protected Boolean returnResidentialAddressLine3;
    protected Boolean returnResidentialCountry;
    protected Boolean returnResidentialCountryDesc;
    protected Boolean returnResidentialState;
    protected Boolean returnResidentialStateDesc;
    protected Boolean returnResidentialZipCode;
    protected Boolean returnResourceClass;
    protected Boolean returnResourceCode;
    protected Boolean returnResourceDesc;
    protected Boolean returnResourceEffectiveDate;
    protected Boolean returnRetirementDate;
    protected Boolean returnSearchStatus;
    protected Boolean returnSeasonalInd;
    protected Boolean returnSeasonalIndDesc;
    protected Boolean returnSecondName;
    protected Boolean returnServiceDate;
    protected Boolean returnSigDateReason1;
    protected Boolean returnSigDateReason2;
    protected Boolean returnSigDateReason3;
    protected Boolean returnSigDateReason4;
    protected Boolean returnSigDateReasonDesc1;
    protected Boolean returnSigDateReasonDesc2;
    protected Boolean returnSigDateReasonDesc3;
    protected Boolean returnSigDateReasonDesc4;
    protected Boolean returnSignificantDate1;
    protected Boolean returnSignificantDate2;
    protected Boolean returnSignificantDate3;
    protected Boolean returnSignificantDate4;
    protected Boolean returnSiteCode;
    protected Boolean returnSiteCodeDesc;
    protected Boolean returnSkillsPassportCode1;
    protected Boolean returnSkillsPassportCode2;
    protected Boolean returnSkillsPassportCode3;
    protected Boolean returnSkillsPassportType1;
    protected Boolean returnSkillsPassportType2;
    protected Boolean returnSkillsPassportType3;
    protected Boolean returnSmokerInd;
    protected Boolean returnSocialInsuranceNumber;
    protected Boolean returnSocialSecurityNoType;
    protected Boolean returnSocialSecurityNoTypeDescription;
    protected Boolean returnSocialSecurityNumber;
    protected Boolean returnStaffCategory;
    protected Boolean returnStaffCategoryDesc;
    protected Boolean returnStdTextRef;
    protected Boolean returnStdTextRefExists;
    protected Boolean returnSupplier;
    protected Boolean returnSupplierName;
    protected Boolean returnSuspensionDate;
    protected Boolean returnTerminationDate;
    protected Boolean returnTerminationReason;
    protected Boolean returnTerminationReasonDesc;
    protected Boolean returnThirdName;
    protected Boolean returnTitle;
    protected Boolean returnTitleDesc;
    protected Boolean returnUnionCode;
    protected Boolean returnUnionCodeDesc;
    protected Boolean returnUserDefContact1;
    protected Boolean returnUserDefContact2;
    protected Boolean returnUserDefContact3;
    protected Boolean returnUserDefContact4;
    protected Boolean returnUserDefContact5;
    protected Boolean returnVeteranStatus;
    protected Boolean returnVeteranStatusDesc;
    protected Boolean returnVisaCode1;
    protected Boolean returnVisaCode1Desc;
    protected Boolean returnVisaCode2;
    protected Boolean returnVisaCode2Desc;
    protected Boolean returnVisaCode3;
    protected Boolean returnVisaCode3Desc;
    protected Boolean returnVisaEffectiveDate1;
    protected Boolean returnVisaEffectiveDate2;
    protected Boolean returnVisaEffectiveDate3;
    protected Boolean returnVisaExpiryDate1;
    protected Boolean returnVisaExpiryDate2;
    protected Boolean returnVisaExpiryDate3;
    protected Boolean returnVisaIssuedBy1;
    protected Boolean returnVisaIssuedBy2;
    protected Boolean returnVisaIssuedBy3;
    protected Boolean returnVisaNumber1;
    protected Boolean returnVisaNumber2;
    protected Boolean returnVisaNumber3;
    protected Boolean returnWorkFacsimileNumber;
    protected Boolean returnWorkGroup;
    protected Boolean returnWorkGroupCrew;
    protected Boolean returnWorkGroupCrewDesc;
    protected Boolean returnWorkGroupDesc;
    protected Boolean returnWorkLocation;
    protected Boolean returnWorkLocationDesc;
    protected Boolean returnWorkMobilePhoneNumber;
    protected Boolean returnWorkOrderPrefix;
    protected Boolean returnWorkOrderPrefixDesc;
    protected Boolean returnWorkTelephoneExtension;
    protected Boolean returnWorkTelephoneNumber;

    /**
     * Gets the value of the returnActualFTEPercent property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnActualFTEPercent() {
        return returnActualFTEPercent;
    }

    /**
     * Sets the value of the returnActualFTEPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnActualFTEPercent(Boolean value) {
        this.returnActualFTEPercent = value;
    }

    /**
     * Gets the value of the returnAuthorityPercent property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAuthorityPercent() {
        return returnAuthorityPercent;
    }

    /**
     * Sets the value of the returnAuthorityPercent property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAuthorityPercent(Boolean value) {
        this.returnAuthorityPercent = value;
    }

    /**
     * Gets the value of the returnAwardCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAwardCode() {
        return returnAwardCode;
    }

    /**
     * Sets the value of the returnAwardCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAwardCode(Boolean value) {
        this.returnAwardCode = value;
    }

    /**
     * Gets the value of the returnAwardCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnAwardCodeDesc() {
        return returnAwardCodeDesc;
    }

    /**
     * Sets the value of the returnAwardCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnAwardCodeDesc(Boolean value) {
        this.returnAwardCodeDesc = value;
    }

    /**
     * Gets the value of the returnBarcode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnBarcode() {
        return returnBarcode;
    }

    /**
     * Sets the value of the returnBarcode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnBarcode(Boolean value) {
        this.returnBarcode = value;
    }

    /**
     * Gets the value of the returnBirthDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnBirthDate() {
        return returnBirthDate;
    }

    /**
     * Sets the value of the returnBirthDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnBirthDate(Boolean value) {
        this.returnBirthDate = value;
    }

    /**
     * Gets the value of the returnBonaFideTermination property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnBonaFideTermination() {
        return returnBonaFideTermination;
    }

    /**
     * Sets the value of the returnBonaFideTermination property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnBonaFideTermination(Boolean value) {
        this.returnBonaFideTermination = value;
    }

    /**
     * Gets the value of the returnCandidateId property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCandidateId() {
        return returnCandidateId;
    }

    /**
     * Sets the value of the returnCandidateId property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCandidateId(Boolean value) {
        this.returnCandidateId = value;
    }

    /**
     * Gets the value of the returnCitizenIndicator property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCitizenIndicator() {
        return returnCitizenIndicator;
    }

    /**
     * Sets the value of the returnCitizenIndicator property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCitizenIndicator(Boolean value) {
        this.returnCitizenIndicator = value;
    }

    /**
     * Gets the value of the returnCitizenIndicatorDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCitizenIndicatorDesc() {
        return returnCitizenIndicatorDesc;
    }

    /**
     * Sets the value of the returnCitizenIndicatorDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCitizenIndicatorDesc(Boolean value) {
        this.returnCitizenIndicatorDesc = value;
    }

    /**
     * Gets the value of the returnCompetencyDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCompetencyDesc() {
        return returnCompetencyDesc;
    }

    /**
     * Sets the value of the returnCompetencyDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCompetencyDesc(Boolean value) {
        this.returnCompetencyDesc = value;
    }

    /**
     * Gets the value of the returnCompetencyLevel property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCompetencyLevel() {
        return returnCompetencyLevel;
    }

    /**
     * Sets the value of the returnCompetencyLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCompetencyLevel(Boolean value) {
        this.returnCompetencyLevel = value;
    }

    /**
     * Gets the value of the returnContractHours property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnContractHours() {
        return returnContractHours;
    }

    /**
     * Sets the value of the returnContractHours property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnContractHours(Boolean value) {
        this.returnContractHours = value;
    }

    /**
     * Gets the value of the returnContractMinutes property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnContractMinutes() {
        return returnContractMinutes;
    }

    /**
     * Sets the value of the returnContractMinutes property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnContractMinutes(Boolean value) {
        this.returnContractMinutes = value;
    }

    /**
     * Gets the value of the returnCopyResAddrPostal property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCopyResAddrPostal() {
        return returnCopyResAddrPostal;
    }

    /**
     * Sets the value of the returnCopyResAddrPostal property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCopyResAddrPostal(Boolean value) {
        this.returnCopyResAddrPostal = value;
    }

    /**
     * Gets the value of the returnCoreEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCoreEmployeeInd() {
        return returnCoreEmployeeInd;
    }

    /**
     * Sets the value of the returnCoreEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCoreEmployeeInd(Boolean value) {
        this.returnCoreEmployeeInd = value;
    }

    /**
     * Gets the value of the returnCountryOfBirth property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCountryOfBirth() {
        return returnCountryOfBirth;
    }

    /**
     * Sets the value of the returnCountryOfBirth property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCountryOfBirth(Boolean value) {
        this.returnCountryOfBirth = value;
    }

    /**
     * Gets the value of the returnCountryOfBirthDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnCountryOfBirthDesc() {
        return returnCountryOfBirthDesc;
    }

    /**
     * Sets the value of the returnCountryOfBirthDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnCountryOfBirthDesc(Boolean value) {
        this.returnCountryOfBirthDesc = value;
    }

    /**
     * Gets the value of the returnDataRefRequired property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDataRefRequired() {
        return returnDataRefRequired;
    }

    /**
     * Sets the value of the returnDataRefRequired property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDataRefRequired(Boolean value) {
        this.returnDataRefRequired = value;
    }

    /**
     * Gets the value of the returnDataReferenceNo property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDataReferenceNo() {
        return returnDataReferenceNo;
    }

    /**
     * Sets the value of the returnDataReferenceNo property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDataReferenceNo(Boolean value) {
        this.returnDataReferenceNo = value;
    }

    /**
     * Gets the value of the returnDeathDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathDate() {
        return returnDeathDate;
    }

    /**
     * Sets the value of the returnDeathDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathDate(Boolean value) {
        this.returnDeathDate = value;
    }

    /**
     * Gets the value of the returnDeathReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathReason() {
        return returnDeathReason;
    }

    /**
     * Sets the value of the returnDeathReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathReason(Boolean value) {
        this.returnDeathReason = value;
    }

    /**
     * Gets the value of the returnDeathReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDeathReasonDesc() {
        return returnDeathReasonDesc;
    }

    /**
     * Sets the value of the returnDeathReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDeathReasonDesc(Boolean value) {
        this.returnDeathReasonDesc = value;
    }

    /**
     * Gets the value of the returnDependants property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDependants() {
        return returnDependants;
    }

    /**
     * Sets the value of the returnDependants property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDependants(Boolean value) {
        this.returnDependants = value;
    }

    /**
     * Gets the value of the returnDisabledInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDisabledInd() {
        return returnDisabledInd;
    }

    /**
     * Sets the value of the returnDisabledInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDisabledInd(Boolean value) {
        this.returnDisabledInd = value;
    }

    /**
     * Gets the value of the returnDuplicateNameInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnDuplicateNameInd() {
        return returnDuplicateNameInd;
    }

    /**
     * Sets the value of the returnDuplicateNameInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnDuplicateNameInd(Boolean value) {
        this.returnDuplicateNameInd = value;
    }

    /**
     * Gets the value of the returnEmailAddress property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmailAddress() {
        return returnEmailAddress;
    }

    /**
     * Sets the value of the returnEmailAddress property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmailAddress(Boolean value) {
        this.returnEmailAddress = value;
    }

    /**
     * Gets the value of the returnEmployee property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployee() {
        return returnEmployee;
    }

    /**
     * Sets the value of the returnEmployee property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployee(Boolean value) {
        this.returnEmployee = value;
    }

    /**
     * Gets the value of the returnEmployeeClass property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeClass() {
        return returnEmployeeClass;
    }

    /**
     * Sets the value of the returnEmployeeClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeClass(Boolean value) {
        this.returnEmployeeClass = value;
    }

    /**
     * Gets the value of the returnEmployeeClassDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeClassDesc() {
        return returnEmployeeClassDesc;
    }

    /**
     * Sets the value of the returnEmployeeClassDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeClassDesc(Boolean value) {
        this.returnEmployeeClassDesc = value;
    }

    /**
     * Gets the value of the returnEmployeeFormattedName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeFormattedName() {
        return returnEmployeeFormattedName;
    }

    /**
     * Sets the value of the returnEmployeeFormattedName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeFormattedName(Boolean value) {
        this.returnEmployeeFormattedName = value;
    }

    /**
     * Gets the value of the returnEmployeeType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeType() {
        return returnEmployeeType;
    }

    /**
     * Sets the value of the returnEmployeeType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeType(Boolean value) {
        this.returnEmployeeType = value;
    }

    /**
     * Gets the value of the returnEmployeeTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEmployeeTypeDesc() {
        return returnEmployeeTypeDesc;
    }

    /**
     * Sets the value of the returnEmployeeTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEmployeeTypeDesc(Boolean value) {
        this.returnEmployeeTypeDesc = value;
    }

    /**
     * Gets the value of the returnEntitleDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEntitleDesc() {
        return returnEntitleDesc;
    }

    /**
     * Sets the value of the returnEntitleDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEntitleDesc(Boolean value) {
        this.returnEntitleDesc = value;
    }

    /**
     * Gets the value of the returnEntitleId property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEntitleId() {
        return returnEntitleId;
    }

    /**
     * Sets the value of the returnEntitleId property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEntitleId(Boolean value) {
        this.returnEntitleId = value;
    }

    /**
     * Gets the value of the returnEssUserInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEssUserInd() {
        return returnEssUserInd;
    }

    /**
     * Sets the value of the returnEssUserInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEssUserInd(Boolean value) {
        this.returnEssUserInd = value;
    }

    /**
     * Gets the value of the returnEthnicity property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEthnicity() {
        return returnEthnicity;
    }

    /**
     * Sets the value of the returnEthnicity property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEthnicity(Boolean value) {
        this.returnEthnicity = value;
    }

    /**
     * Gets the value of the returnEthnicityDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnEthnicityDesc() {
        return returnEthnicityDesc;
    }

    /**
     * Sets the value of the returnEthnicityDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnEthnicityDesc(Boolean value) {
        this.returnEthnicityDesc = value;
    }

    /**
     * Gets the value of the returnExcludeTalentExtract property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnExcludeTalentExtract() {
        return returnExcludeTalentExtract;
    }

    /**
     * Sets the value of the returnExcludeTalentExtract property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnExcludeTalentExtract(Boolean value) {
        this.returnExcludeTalentExtract = value;
    }

    /**
     * Gets the value of the returnFirstName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnFirstName() {
        return returnFirstName;
    }

    /**
     * Sets the value of the returnFirstName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnFirstName(Boolean value) {
        this.returnFirstName = value;
    }

    /**
     * Gets the value of the returnFixedAssetsDistrict property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnFixedAssetsDistrict() {
        return returnFixedAssetsDistrict;
    }

    /**
     * Sets the value of the returnFixedAssetsDistrict property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnFixedAssetsDistrict(Boolean value) {
        this.returnFixedAssetsDistrict = value;
    }

    /**
     * Gets the value of the returnFixedAssetsDistrictDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnFixedAssetsDistrictDesc() {
        return returnFixedAssetsDistrictDesc;
    }

    /**
     * Sets the value of the returnFixedAssetsDistrictDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnFixedAssetsDistrictDesc(Boolean value) {
        this.returnFixedAssetsDistrictDesc = value;
    }

    /**
     * Gets the value of the returnGender property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnGender() {
        return returnGender;
    }

    /**
     * Sets the value of the returnGender property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnGender(Boolean value) {
        this.returnGender = value;
    }

    /**
     * Gets the value of the returnGenderDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnGenderDesc() {
        return returnGenderDesc;
    }

    /**
     * Sets the value of the returnGenderDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnGenderDesc(Boolean value) {
        this.returnGenderDesc = value;
    }

    /**
     * Gets the value of the returnGlobalProfile property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnGlobalProfile() {
        return returnGlobalProfile;
    }

    /**
     * Sets the value of the returnGlobalProfile property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnGlobalProfile(Boolean value) {
        this.returnGlobalProfile = value;
    }

    /**
     * Gets the value of the returnHealthPlan property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHealthPlan() {
        return returnHealthPlan;
    }

    /**
     * Sets the value of the returnHealthPlan property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHealthPlan(Boolean value) {
        this.returnHealthPlan = value;
    }

    /**
     * Gets the value of the returnHireDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHireDate() {
        return returnHireDate;
    }

    /**
     * Sets the value of the returnHireDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHireDate(Boolean value) {
        this.returnHireDate = value;
    }

    /**
     * Gets the value of the returnHomeFacsimileNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHomeFacsimileNumber() {
        return returnHomeFacsimileNumber;
    }

    /**
     * Sets the value of the returnHomeFacsimileNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHomeFacsimileNumber(Boolean value) {
        this.returnHomeFacsimileNumber = value;
    }

    /**
     * Gets the value of the returnHomeMobilePhoneNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHomeMobilePhoneNumber() {
        return returnHomeMobilePhoneNumber;
    }

    /**
     * Sets the value of the returnHomeMobilePhoneNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHomeMobilePhoneNumber(Boolean value) {
        this.returnHomeMobilePhoneNumber = value;
    }

    /**
     * Gets the value of the returnHomeTelephoneNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnHomeTelephoneNumber() {
        return returnHomeTelephoneNumber;
    }

    /**
     * Sets the value of the returnHomeTelephoneNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnHomeTelephoneNumber(Boolean value) {
        this.returnHomeTelephoneNumber = value;
    }

    /**
     * Gets the value of the returnJobClassLevel property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnJobClassLevel() {
        return returnJobClassLevel;
    }

    /**
     * Sets the value of the returnJobClassLevel property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnJobClassLevel(Boolean value) {
        this.returnJobClassLevel = value;
    }

    /**
     * Gets the value of the returnJobClassLevelDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnJobClassLevelDesc() {
        return returnJobClassLevelDesc;
    }

    /**
     * Sets the value of the returnJobClassLevelDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnJobClassLevelDesc(Boolean value) {
        this.returnJobClassLevelDesc = value;
    }

    /**
     * Gets the value of the returnLanguageCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLanguageCode() {
        return returnLanguageCode;
    }

    /**
     * Sets the value of the returnLanguageCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLanguageCode(Boolean value) {
        this.returnLanguageCode = value;
    }

    /**
     * Gets the value of the returnLanguageCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLanguageCodeDesc() {
        return returnLanguageCodeDesc;
    }

    /**
     * Sets the value of the returnLanguageCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLanguageCodeDesc(Boolean value) {
        this.returnLanguageCodeDesc = value;
    }

    /**
     * Gets the value of the returnLastName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLastName() {
        return returnLastName;
    }

    /**
     * Sets the value of the returnLastName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLastName(Boolean value) {
        this.returnLastName = value;
    }

    /**
     * Gets the value of the returnLeaveForecastDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLeaveForecastDate() {
        return returnLeaveForecastDate;
    }

    /**
     * Sets the value of the returnLeaveForecastDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLeaveForecastDate(Boolean value) {
        this.returnLeaveForecastDate = value;
    }

    /**
     * Gets the value of the returnLegalRepIndicator property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLegalRepIndicator() {
        return returnLegalRepIndicator;
    }

    /**
     * Sets the value of the returnLegalRepIndicator property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLegalRepIndicator(Boolean value) {
        this.returnLegalRepIndicator = value;
    }

    /**
     * Gets the value of the returnLegalRepName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnLegalRepName() {
        return returnLegalRepName;
    }

    /**
     * Sets the value of the returnLegalRepName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnLegalRepName(Boolean value) {
        this.returnLegalRepName = value;
    }

    /**
     * Gets the value of the returnMaritalStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMaritalStatus() {
        return returnMaritalStatus;
    }

    /**
     * Sets the value of the returnMaritalStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMaritalStatus(Boolean value) {
        this.returnMaritalStatus = value;
    }

    /**
     * Gets the value of the returnMaritalStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMaritalStatusDesc() {
        return returnMaritalStatusDesc;
    }

    /**
     * Sets the value of the returnMaritalStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMaritalStatusDesc(Boolean value) {
        this.returnMaritalStatusDesc = value;
    }

    /**
     * Gets the value of the returnMessagePreference property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMessagePreference() {
        return returnMessagePreference;
    }

    /**
     * Sets the value of the returnMessagePreference property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMessagePreference(Boolean value) {
        this.returnMessagePreference = value;
    }

    /**
     * Gets the value of the returnMessagePreferenceDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnMessagePreferenceDesc() {
        return returnMessagePreferenceDesc;
    }

    /**
     * Sets the value of the returnMessagePreferenceDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnMessagePreferenceDesc(Boolean value) {
        this.returnMessagePreferenceDesc = value;
    }

    /**
     * Gets the value of the returnNationalIdCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdCode1() {
        return returnNationalIdCode1;
    }

    /**
     * Sets the value of the returnNationalIdCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdCode1(Boolean value) {
        this.returnNationalIdCode1 = value;
    }

    /**
     * Gets the value of the returnNationalIdCode2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdCode2() {
        return returnNationalIdCode2;
    }

    /**
     * Sets the value of the returnNationalIdCode2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdCode2(Boolean value) {
        this.returnNationalIdCode2 = value;
    }

    /**
     * Gets the value of the returnNationalIdCode3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdCode3() {
        return returnNationalIdCode3;
    }

    /**
     * Sets the value of the returnNationalIdCode3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdCode3(Boolean value) {
        this.returnNationalIdCode3 = value;
    }

    /**
     * Gets the value of the returnNationalIdCode4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdCode4() {
        return returnNationalIdCode4;
    }

    /**
     * Sets the value of the returnNationalIdCode4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdCode4(Boolean value) {
        this.returnNationalIdCode4 = value;
    }

    /**
     * Gets the value of the returnNationalIdCode5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdCode5() {
        return returnNationalIdCode5;
    }

    /**
     * Sets the value of the returnNationalIdCode5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdCode5(Boolean value) {
        this.returnNationalIdCode5 = value;
    }

    /**
     * Gets the value of the returnNationalIdType1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType1() {
        return returnNationalIdType1;
    }

    /**
     * Sets the value of the returnNationalIdType1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType1(Boolean value) {
        this.returnNationalIdType1 = value;
    }

    /**
     * Gets the value of the returnNationalIdType1Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType1Desc() {
        return returnNationalIdType1Desc;
    }

    /**
     * Sets the value of the returnNationalIdType1Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType1Desc(Boolean value) {
        this.returnNationalIdType1Desc = value;
    }

    /**
     * Gets the value of the returnNationalIdType2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType2() {
        return returnNationalIdType2;
    }

    /**
     * Sets the value of the returnNationalIdType2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType2(Boolean value) {
        this.returnNationalIdType2 = value;
    }

    /**
     * Gets the value of the returnNationalIdType2Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType2Desc() {
        return returnNationalIdType2Desc;
    }

    /**
     * Sets the value of the returnNationalIdType2Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType2Desc(Boolean value) {
        this.returnNationalIdType2Desc = value;
    }

    /**
     * Gets the value of the returnNationalIdType3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType3() {
        return returnNationalIdType3;
    }

    /**
     * Sets the value of the returnNationalIdType3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType3(Boolean value) {
        this.returnNationalIdType3 = value;
    }

    /**
     * Gets the value of the returnNationalIdType3Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType3Desc() {
        return returnNationalIdType3Desc;
    }

    /**
     * Sets the value of the returnNationalIdType3Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType3Desc(Boolean value) {
        this.returnNationalIdType3Desc = value;
    }

    /**
     * Gets the value of the returnNationalIdType4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType4() {
        return returnNationalIdType4;
    }

    /**
     * Sets the value of the returnNationalIdType4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType4(Boolean value) {
        this.returnNationalIdType4 = value;
    }

    /**
     * Gets the value of the returnNationalIdType4Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType4Desc() {
        return returnNationalIdType4Desc;
    }

    /**
     * Sets the value of the returnNationalIdType4Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType4Desc(Boolean value) {
        this.returnNationalIdType4Desc = value;
    }

    /**
     * Gets the value of the returnNationalIdType5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType5() {
        return returnNationalIdType5;
    }

    /**
     * Sets the value of the returnNationalIdType5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType5(Boolean value) {
        this.returnNationalIdType5 = value;
    }

    /**
     * Gets the value of the returnNationalIdType5Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalIdType5Desc() {
        return returnNationalIdType5Desc;
    }

    /**
     * Sets the value of the returnNationalIdType5Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalIdType5Desc(Boolean value) {
        this.returnNationalIdType5Desc = value;
    }

    /**
     * Gets the value of the returnNationality property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationality() {
        return returnNationality;
    }

    /**
     * Sets the value of the returnNationality property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationality(Boolean value) {
        this.returnNationality = value;
    }

    /**
     * Gets the value of the returnNationalityDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNationalityDesc() {
        return returnNationalityDesc;
    }

    /**
     * Sets the value of the returnNationalityDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNationalityDesc(Boolean value) {
        this.returnNationalityDesc = value;
    }

    /**
     * Gets the value of the returnNotifyEDIMsgRecieved property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnNotifyEDIMsgRecieved() {
        return returnNotifyEDIMsgRecieved;
    }

    /**
     * Sets the value of the returnNotifyEDIMsgRecieved property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnNotifyEDIMsgRecieved(Boolean value) {
        this.returnNotifyEDIMsgRecieved = value;
    }

    /**
     * Gets the value of the returnPassportExpiryDate1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportExpiryDate1() {
        return returnPassportExpiryDate1;
    }

    /**
     * Sets the value of the returnPassportExpiryDate1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportExpiryDate1(Boolean value) {
        this.returnPassportExpiryDate1 = value;
    }

    /**
     * Gets the value of the returnPassportExpiryDate2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportExpiryDate2() {
        return returnPassportExpiryDate2;
    }

    /**
     * Sets the value of the returnPassportExpiryDate2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportExpiryDate2(Boolean value) {
        this.returnPassportExpiryDate2 = value;
    }

    /**
     * Gets the value of the returnPassportExpiryDate3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportExpiryDate3() {
        return returnPassportExpiryDate3;
    }

    /**
     * Sets the value of the returnPassportExpiryDate3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportExpiryDate3(Boolean value) {
        this.returnPassportExpiryDate3 = value;
    }

    /**
     * Gets the value of the returnPassportIssuedBy1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportIssuedBy1() {
        return returnPassportIssuedBy1;
    }

    /**
     * Sets the value of the returnPassportIssuedBy1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportIssuedBy1(Boolean value) {
        this.returnPassportIssuedBy1 = value;
    }

    /**
     * Gets the value of the returnPassportIssuedBy2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportIssuedBy2() {
        return returnPassportIssuedBy2;
    }

    /**
     * Sets the value of the returnPassportIssuedBy2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportIssuedBy2(Boolean value) {
        this.returnPassportIssuedBy2 = value;
    }

    /**
     * Gets the value of the returnPassportIssuedBy3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportIssuedBy3() {
        return returnPassportIssuedBy3;
    }

    /**
     * Sets the value of the returnPassportIssuedBy3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportIssuedBy3(Boolean value) {
        this.returnPassportIssuedBy3 = value;
    }

    /**
     * Gets the value of the returnPassportNumber1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportNumber1() {
        return returnPassportNumber1;
    }

    /**
     * Sets the value of the returnPassportNumber1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportNumber1(Boolean value) {
        this.returnPassportNumber1 = value;
    }

    /**
     * Gets the value of the returnPassportNumber2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportNumber2() {
        return returnPassportNumber2;
    }

    /**
     * Sets the value of the returnPassportNumber2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportNumber2(Boolean value) {
        this.returnPassportNumber2 = value;
    }

    /**
     * Gets the value of the returnPassportNumber3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPassportNumber3() {
        return returnPassportNumber3;
    }

    /**
     * Sets the value of the returnPassportNumber3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPassportNumber3(Boolean value) {
        this.returnPassportNumber3 = value;
    }

    /**
     * Gets the value of the returnPaygroup property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPaygroup() {
        return returnPaygroup;
    }

    /**
     * Sets the value of the returnPaygroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPaygroup(Boolean value) {
        this.returnPaygroup = value;
    }

    /**
     * Gets the value of the returnPaygroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPaygroupDesc() {
        return returnPaygroupDesc;
    }

    /**
     * Sets the value of the returnPaygroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPaygroupDesc(Boolean value) {
        this.returnPaygroupDesc = value;
    }

    /**
     * Gets the value of the returnPayrollEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPayrollEmployeeInd() {
        return returnPayrollEmployeeInd;
    }

    /**
     * Sets the value of the returnPayrollEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPayrollEmployeeInd(Boolean value) {
        this.returnPayrollEmployeeInd = value;
    }

    /**
     * Gets the value of the returnPersEmpStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersEmpStatus() {
        return returnPersEmpStatus;
    }

    /**
     * Sets the value of the returnPersEmpStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersEmpStatus(Boolean value) {
        this.returnPersEmpStatus = value;
    }

    /**
     * Gets the value of the returnPersEmpStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersEmpStatusDesc() {
        return returnPersEmpStatusDesc;
    }

    /**
     * Sets the value of the returnPersEmpStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersEmpStatusDesc(Boolean value) {
        this.returnPersEmpStatusDesc = value;
    }

    /**
     * Gets the value of the returnPersonType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonType() {
        return returnPersonType;
    }

    /**
     * Sets the value of the returnPersonType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonType(Boolean value) {
        this.returnPersonType = value;
    }

    /**
     * Gets the value of the returnPersonTypeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonTypeDesc() {
        return returnPersonTypeDesc;
    }

    /**
     * Sets the value of the returnPersonTypeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonTypeDesc(Boolean value) {
        this.returnPersonTypeDesc = value;
    }

    /**
     * Gets the value of the returnPersonalEmail property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonalEmail() {
        return returnPersonalEmail;
    }

    /**
     * Sets the value of the returnPersonalEmail property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonalEmail(Boolean value) {
        this.returnPersonalEmail = value;
    }

    /**
     * Gets the value of the returnPersonnelClass1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass1() {
        return returnPersonnelClass1;
    }

    /**
     * Sets the value of the returnPersonnelClass1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass1(Boolean value) {
        this.returnPersonnelClass1 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass10 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass10() {
        return returnPersonnelClass10;
    }

    /**
     * Sets the value of the returnPersonnelClass10 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass10(Boolean value) {
        this.returnPersonnelClass10 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass10Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass10Desc() {
        return returnPersonnelClass10Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass10Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass10Desc(Boolean value) {
        this.returnPersonnelClass10Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass10Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass10Name() {
        return returnPersonnelClass10Name;
    }

    /**
     * Sets the value of the returnPersonnelClass10Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass10Name(Boolean value) {
        this.returnPersonnelClass10Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass1Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass1Desc() {
        return returnPersonnelClass1Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass1Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass1Desc(Boolean value) {
        this.returnPersonnelClass1Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass1Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass1Name() {
        return returnPersonnelClass1Name;
    }

    /**
     * Sets the value of the returnPersonnelClass1Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass1Name(Boolean value) {
        this.returnPersonnelClass1Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass2() {
        return returnPersonnelClass2;
    }

    /**
     * Sets the value of the returnPersonnelClass2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass2(Boolean value) {
        this.returnPersonnelClass2 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass2Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass2Desc() {
        return returnPersonnelClass2Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass2Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass2Desc(Boolean value) {
        this.returnPersonnelClass2Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass2Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass2Name() {
        return returnPersonnelClass2Name;
    }

    /**
     * Sets the value of the returnPersonnelClass2Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass2Name(Boolean value) {
        this.returnPersonnelClass2Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass3() {
        return returnPersonnelClass3;
    }

    /**
     * Sets the value of the returnPersonnelClass3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass3(Boolean value) {
        this.returnPersonnelClass3 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass3Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass3Desc() {
        return returnPersonnelClass3Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass3Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass3Desc(Boolean value) {
        this.returnPersonnelClass3Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass3Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass3Name() {
        return returnPersonnelClass3Name;
    }

    /**
     * Sets the value of the returnPersonnelClass3Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass3Name(Boolean value) {
        this.returnPersonnelClass3Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass4() {
        return returnPersonnelClass4;
    }

    /**
     * Sets the value of the returnPersonnelClass4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass4(Boolean value) {
        this.returnPersonnelClass4 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass4Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass4Desc() {
        return returnPersonnelClass4Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass4Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass4Desc(Boolean value) {
        this.returnPersonnelClass4Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass4Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass4Name() {
        return returnPersonnelClass4Name;
    }

    /**
     * Sets the value of the returnPersonnelClass4Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass4Name(Boolean value) {
        this.returnPersonnelClass4Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass5() {
        return returnPersonnelClass5;
    }

    /**
     * Sets the value of the returnPersonnelClass5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass5(Boolean value) {
        this.returnPersonnelClass5 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass5Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass5Desc() {
        return returnPersonnelClass5Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass5Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass5Desc(Boolean value) {
        this.returnPersonnelClass5Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass5Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass5Name() {
        return returnPersonnelClass5Name;
    }

    /**
     * Sets the value of the returnPersonnelClass5Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass5Name(Boolean value) {
        this.returnPersonnelClass5Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass6 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass6() {
        return returnPersonnelClass6;
    }

    /**
     * Sets the value of the returnPersonnelClass6 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass6(Boolean value) {
        this.returnPersonnelClass6 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass6Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass6Desc() {
        return returnPersonnelClass6Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass6Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass6Desc(Boolean value) {
        this.returnPersonnelClass6Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass6Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass6Name() {
        return returnPersonnelClass6Name;
    }

    /**
     * Sets the value of the returnPersonnelClass6Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass6Name(Boolean value) {
        this.returnPersonnelClass6Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass7 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass7() {
        return returnPersonnelClass7;
    }

    /**
     * Sets the value of the returnPersonnelClass7 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass7(Boolean value) {
        this.returnPersonnelClass7 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass7Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass7Desc() {
        return returnPersonnelClass7Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass7Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass7Desc(Boolean value) {
        this.returnPersonnelClass7Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass7Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass7Name() {
        return returnPersonnelClass7Name;
    }

    /**
     * Sets the value of the returnPersonnelClass7Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass7Name(Boolean value) {
        this.returnPersonnelClass7Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass8 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass8() {
        return returnPersonnelClass8;
    }

    /**
     * Sets the value of the returnPersonnelClass8 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass8(Boolean value) {
        this.returnPersonnelClass8 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass8Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass8Desc() {
        return returnPersonnelClass8Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass8Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass8Desc(Boolean value) {
        this.returnPersonnelClass8Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass8Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass8Name() {
        return returnPersonnelClass8Name;
    }

    /**
     * Sets the value of the returnPersonnelClass8Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass8Name(Boolean value) {
        this.returnPersonnelClass8Name = value;
    }

    /**
     * Gets the value of the returnPersonnelClass9 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass9() {
        return returnPersonnelClass9;
    }

    /**
     * Sets the value of the returnPersonnelClass9 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass9(Boolean value) {
        this.returnPersonnelClass9 = value;
    }

    /**
     * Gets the value of the returnPersonnelClass9Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass9Desc() {
        return returnPersonnelClass9Desc;
    }

    /**
     * Sets the value of the returnPersonnelClass9Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass9Desc(Boolean value) {
        this.returnPersonnelClass9Desc = value;
    }

    /**
     * Gets the value of the returnPersonnelClass9Name property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelClass9Name() {
        return returnPersonnelClass9Name;
    }

    /**
     * Sets the value of the returnPersonnelClass9Name property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelClass9Name(Boolean value) {
        this.returnPersonnelClass9Name = value;
    }

    /**
     * Gets the value of the returnPersonnelEmployeeInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelEmployeeInd() {
        return returnPersonnelEmployeeInd;
    }

    /**
     * Sets the value of the returnPersonnelEmployeeInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelEmployeeInd(Boolean value) {
        this.returnPersonnelEmployeeInd = value;
    }

    /**
     * Gets the value of the returnPersonnelGroup property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelGroup() {
        return returnPersonnelGroup;
    }

    /**
     * Sets the value of the returnPersonnelGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelGroup(Boolean value) {
        this.returnPersonnelGroup = value;
    }

    /**
     * Gets the value of the returnPersonnelGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelGroupDesc() {
        return returnPersonnelGroupDesc;
    }

    /**
     * Sets the value of the returnPersonnelGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelGroupDesc(Boolean value) {
        this.returnPersonnelGroupDesc = value;
    }

    /**
     * Gets the value of the returnPersonnelStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelStatus() {
        return returnPersonnelStatus;
    }

    /**
     * Sets the value of the returnPersonnelStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelStatus(Boolean value) {
        this.returnPersonnelStatus = value;
    }

    /**
     * Gets the value of the returnPersonnelStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPersonnelStatusDesc() {
        return returnPersonnelStatusDesc;
    }

    /**
     * Sets the value of the returnPersonnelStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPersonnelStatusDesc(Boolean value) {
        this.returnPersonnelStatusDesc = value;
    }

    /**
     * Gets the value of the returnPhotoPathname property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhotoPathname() {
        return returnPhotoPathname;
    }

    /**
     * Sets the value of the returnPhotoPathname property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhotoPathname(Boolean value) {
        this.returnPhotoPathname = value;
    }

    /**
     * Gets the value of the returnPhysicalLocReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocReason() {
        return returnPhysicalLocReason;
    }

    /**
     * Sets the value of the returnPhysicalLocReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocReason(Boolean value) {
        this.returnPhysicalLocReason = value;
    }

    /**
     * Gets the value of the returnPhysicalLocReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocReasonDesc() {
        return returnPhysicalLocReasonDesc;
    }

    /**
     * Sets the value of the returnPhysicalLocReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocReasonDesc(Boolean value) {
        this.returnPhysicalLocReasonDesc = value;
    }

    /**
     * Gets the value of the returnPhysicalLocation property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocation() {
        return returnPhysicalLocation;
    }

    /**
     * Sets the value of the returnPhysicalLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocation(Boolean value) {
        this.returnPhysicalLocation = value;
    }

    /**
     * Gets the value of the returnPhysicalLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPhysicalLocationDesc() {
        return returnPhysicalLocationDesc;
    }

    /**
     * Sets the value of the returnPhysicalLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPhysicalLocationDesc(Boolean value) {
        this.returnPhysicalLocationDesc = value;
    }

    /**
     * Gets the value of the returnPosition property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPosition() {
        return returnPosition;
    }

    /**
     * Sets the value of the returnPosition property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPosition(Boolean value) {
        this.returnPosition = value;
    }

    /**
     * Gets the value of the returnPositionClassUDef1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionClassUDef1() {
        return returnPositionClassUDef1;
    }

    /**
     * Sets the value of the returnPositionClassUDef1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionClassUDef1(Boolean value) {
        this.returnPositionClassUDef1 = value;
    }

    /**
     * Gets the value of the returnPositionClassUDef2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionClassUDef2() {
        return returnPositionClassUDef2;
    }

    /**
     * Sets the value of the returnPositionClassUDef2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionClassUDef2(Boolean value) {
        this.returnPositionClassUDef2 = value;
    }

    /**
     * Gets the value of the returnPositionClassUDefName1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionClassUDefName1() {
        return returnPositionClassUDefName1;
    }

    /**
     * Sets the value of the returnPositionClassUDefName1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionClassUDefName1(Boolean value) {
        this.returnPositionClassUDefName1 = value;
    }

    /**
     * Gets the value of the returnPositionClassUDefName2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionClassUDefName2() {
        return returnPositionClassUDefName2;
    }

    /**
     * Sets the value of the returnPositionClassUDefName2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionClassUDefName2(Boolean value) {
        this.returnPositionClassUDefName2 = value;
    }

    /**
     * Gets the value of the returnPositionDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionDesc() {
        return returnPositionDesc;
    }

    /**
     * Sets the value of the returnPositionDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionDesc(Boolean value) {
        this.returnPositionDesc = value;
    }

    /**
     * Gets the value of the returnPositionReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionReason() {
        return returnPositionReason;
    }

    /**
     * Sets the value of the returnPositionReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionReason(Boolean value) {
        this.returnPositionReason = value;
    }

    /**
     * Gets the value of the returnPositionReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionReasonDesc() {
        return returnPositionReasonDesc;
    }

    /**
     * Sets the value of the returnPositionReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionReasonDesc(Boolean value) {
        this.returnPositionReasonDesc = value;
    }

    /**
     * Gets the value of the returnPositionStartDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPositionStartDate() {
        return returnPositionStartDate;
    }

    /**
     * Sets the value of the returnPositionStartDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPositionStartDate(Boolean value) {
        this.returnPositionStartDate = value;
    }

    /**
     * Gets the value of the returnPostalAddressLine1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalAddressLine1() {
        return returnPostalAddressLine1;
    }

    /**
     * Sets the value of the returnPostalAddressLine1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalAddressLine1(Boolean value) {
        this.returnPostalAddressLine1 = value;
    }

    /**
     * Gets the value of the returnPostalAddressLine2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalAddressLine2() {
        return returnPostalAddressLine2;
    }

    /**
     * Sets the value of the returnPostalAddressLine2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalAddressLine2(Boolean value) {
        this.returnPostalAddressLine2 = value;
    }

    /**
     * Gets the value of the returnPostalAddressLine3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalAddressLine3() {
        return returnPostalAddressLine3;
    }

    /**
     * Sets the value of the returnPostalAddressLine3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalAddressLine3(Boolean value) {
        this.returnPostalAddressLine3 = value;
    }

    /**
     * Gets the value of the returnPostalCountry property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalCountry() {
        return returnPostalCountry;
    }

    /**
     * Sets the value of the returnPostalCountry property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalCountry(Boolean value) {
        this.returnPostalCountry = value;
    }

    /**
     * Gets the value of the returnPostalCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalCountryDesc() {
        return returnPostalCountryDesc;
    }

    /**
     * Sets the value of the returnPostalCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalCountryDesc(Boolean value) {
        this.returnPostalCountryDesc = value;
    }

    /**
     * Gets the value of the returnPostalState property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalState() {
        return returnPostalState;
    }

    /**
     * Sets the value of the returnPostalState property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalState(Boolean value) {
        this.returnPostalState = value;
    }

    /**
     * Gets the value of the returnPostalStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalStateDesc() {
        return returnPostalStateDesc;
    }

    /**
     * Sets the value of the returnPostalStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalStateDesc(Boolean value) {
        this.returnPostalStateDesc = value;
    }

    /**
     * Gets the value of the returnPostalZipCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPostalZipCode() {
        return returnPostalZipCode;
    }

    /**
     * Sets the value of the returnPostalZipCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPostalZipCode(Boolean value) {
        this.returnPostalZipCode = value;
    }

    /**
     * Gets the value of the returnPreferredName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreferredName() {
        return returnPreferredName;
    }

    /**
     * Sets the value of the returnPreferredName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreferredName(Boolean value) {
        this.returnPreferredName = value;
    }

    /**
     * Gets the value of the returnPreviousEmployeeId property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreviousEmployeeId() {
        return returnPreviousEmployeeId;
    }

    /**
     * Sets the value of the returnPreviousEmployeeId property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreviousEmployeeId(Boolean value) {
        this.returnPreviousEmployeeId = value;
    }

    /**
     * Gets the value of the returnPreviousFirstName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreviousFirstName() {
        return returnPreviousFirstName;
    }

    /**
     * Sets the value of the returnPreviousFirstName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreviousFirstName(Boolean value) {
        this.returnPreviousFirstName = value;
    }

    /**
     * Gets the value of the returnPreviousLastName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreviousLastName() {
        return returnPreviousLastName;
    }

    /**
     * Sets the value of the returnPreviousLastName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreviousLastName(Boolean value) {
        this.returnPreviousLastName = value;
    }

    /**
     * Gets the value of the returnPreviousSecondName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreviousSecondName() {
        return returnPreviousSecondName;
    }

    /**
     * Sets the value of the returnPreviousSecondName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreviousSecondName(Boolean value) {
        this.returnPreviousSecondName = value;
    }

    /**
     * Gets the value of the returnPreviousThirdName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPreviousThirdName() {
        return returnPreviousThirdName;
    }

    /**
     * Sets the value of the returnPreviousThirdName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPreviousThirdName(Boolean value) {
        this.returnPreviousThirdName = value;
    }

    /**
     * Gets the value of the returnPrimRepCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrimRepCode() {
        return returnPrimRepCode;
    }

    /**
     * Sets the value of the returnPrimRepCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrimRepCode(Boolean value) {
        this.returnPrimRepCode = value;
    }

    /**
     * Gets the value of the returnPrimRepCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrimRepCodeDesc() {
        return returnPrimRepCodeDesc;
    }

    /**
     * Sets the value of the returnPrimRepCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrimRepCodeDesc(Boolean value) {
        this.returnPrimRepCodeDesc = value;
    }

    /**
     * Gets the value of the returnPrinterCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode1() {
        return returnPrinterCode1;
    }

    /**
     * Sets the value of the returnPrinterCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode1(Boolean value) {
        this.returnPrinterCode1 = value;
    }

    /**
     * Gets the value of the returnPrinterCode2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode2() {
        return returnPrinterCode2;
    }

    /**
     * Sets the value of the returnPrinterCode2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode2(Boolean value) {
        this.returnPrinterCode2 = value;
    }

    /**
     * Gets the value of the returnPrinterCode3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode3() {
        return returnPrinterCode3;
    }

    /**
     * Sets the value of the returnPrinterCode3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode3(Boolean value) {
        this.returnPrinterCode3 = value;
    }

    /**
     * Gets the value of the returnPrinterCode4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode4() {
        return returnPrinterCode4;
    }

    /**
     * Sets the value of the returnPrinterCode4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode4(Boolean value) {
        this.returnPrinterCode4 = value;
    }

    /**
     * Gets the value of the returnPrinterCode5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterCode5() {
        return returnPrinterCode5;
    }

    /**
     * Sets the value of the returnPrinterCode5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterCode5(Boolean value) {
        this.returnPrinterCode5 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc1() {
        return returnPrinterDesc1;
    }

    /**
     * Sets the value of the returnPrinterDesc1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc1(Boolean value) {
        this.returnPrinterDesc1 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc2() {
        return returnPrinterDesc2;
    }

    /**
     * Sets the value of the returnPrinterDesc2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc2(Boolean value) {
        this.returnPrinterDesc2 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc3() {
        return returnPrinterDesc3;
    }

    /**
     * Sets the value of the returnPrinterDesc3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc3(Boolean value) {
        this.returnPrinterDesc3 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc4() {
        return returnPrinterDesc4;
    }

    /**
     * Sets the value of the returnPrinterDesc4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc4(Boolean value) {
        this.returnPrinterDesc4 = value;
    }

    /**
     * Gets the value of the returnPrinterDesc5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterDesc5() {
        return returnPrinterDesc5;
    }

    /**
     * Sets the value of the returnPrinterDesc5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterDesc5(Boolean value) {
        this.returnPrinterDesc5 = value;
    }

    /**
     * Gets the value of the returnPrinterName1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterName1() {
        return returnPrinterName1;
    }

    /**
     * Sets the value of the returnPrinterName1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterName1(Boolean value) {
        this.returnPrinterName1 = value;
    }

    /**
     * Gets the value of the returnPrinterName2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterName2() {
        return returnPrinterName2;
    }

    /**
     * Sets the value of the returnPrinterName2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterName2(Boolean value) {
        this.returnPrinterName2 = value;
    }

    /**
     * Gets the value of the returnPrinterName3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterName3() {
        return returnPrinterName3;
    }

    /**
     * Sets the value of the returnPrinterName3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterName3(Boolean value) {
        this.returnPrinterName3 = value;
    }

    /**
     * Gets the value of the returnPrinterName4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterName4() {
        return returnPrinterName4;
    }

    /**
     * Sets the value of the returnPrinterName4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterName4(Boolean value) {
        this.returnPrinterName4 = value;
    }

    /**
     * Gets the value of the returnPrinterName5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnPrinterName5() {
        return returnPrinterName5;
    }

    /**
     * Sets the value of the returnPrinterName5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnPrinterName5(Boolean value) {
        this.returnPrinterName5 = value;
    }

    /**
     * Gets the value of the returnProfessionalServiceDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnProfessionalServiceDate() {
        return returnProfessionalServiceDate;
    }

    /**
     * Sets the value of the returnProfessionalServiceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnProfessionalServiceDate(Boolean value) {
        this.returnProfessionalServiceDate = value;
    }

    /**
     * Gets the value of the returnRehireCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnRehireCode() {
        return returnRehireCode;
    }

    /**
     * Sets the value of the returnRehireCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnRehireCode(Boolean value) {
        this.returnRehireCode = value;
    }

    /**
     * Gets the value of the returnRehireCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnRehireCodeDesc() {
        return returnRehireCodeDesc;
    }

    /**
     * Sets the value of the returnRehireCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnRehireCodeDesc(Boolean value) {
        this.returnRehireCodeDesc = value;
    }

    /**
     * Gets the value of the returnReinstatementDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnReinstatementDate() {
        return returnReinstatementDate;
    }

    /**
     * Sets the value of the returnReinstatementDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnReinstatementDate(Boolean value) {
        this.returnReinstatementDate = value;
    }

    /**
     * Gets the value of the returnResidentialAddressEffDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialAddressEffDate() {
        return returnResidentialAddressEffDate;
    }

    /**
     * Sets the value of the returnResidentialAddressEffDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialAddressEffDate(Boolean value) {
        this.returnResidentialAddressEffDate = value;
    }

    /**
     * Gets the value of the returnResidentialAddressLine1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialAddressLine1() {
        return returnResidentialAddressLine1;
    }

    /**
     * Sets the value of the returnResidentialAddressLine1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialAddressLine1(Boolean value) {
        this.returnResidentialAddressLine1 = value;
    }

    /**
     * Gets the value of the returnResidentialAddressLine2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialAddressLine2() {
        return returnResidentialAddressLine2;
    }

    /**
     * Sets the value of the returnResidentialAddressLine2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialAddressLine2(Boolean value) {
        this.returnResidentialAddressLine2 = value;
    }

    /**
     * Gets the value of the returnResidentialAddressLine3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialAddressLine3() {
        return returnResidentialAddressLine3;
    }

    /**
     * Sets the value of the returnResidentialAddressLine3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialAddressLine3(Boolean value) {
        this.returnResidentialAddressLine3 = value;
    }

    /**
     * Gets the value of the returnResidentialCountry property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialCountry() {
        return returnResidentialCountry;
    }

    /**
     * Sets the value of the returnResidentialCountry property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialCountry(Boolean value) {
        this.returnResidentialCountry = value;
    }

    /**
     * Gets the value of the returnResidentialCountryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialCountryDesc() {
        return returnResidentialCountryDesc;
    }

    /**
     * Sets the value of the returnResidentialCountryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialCountryDesc(Boolean value) {
        this.returnResidentialCountryDesc = value;
    }

    /**
     * Gets the value of the returnResidentialState property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialState() {
        return returnResidentialState;
    }

    /**
     * Sets the value of the returnResidentialState property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialState(Boolean value) {
        this.returnResidentialState = value;
    }

    /**
     * Gets the value of the returnResidentialStateDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialStateDesc() {
        return returnResidentialStateDesc;
    }

    /**
     * Sets the value of the returnResidentialStateDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialStateDesc(Boolean value) {
        this.returnResidentialStateDesc = value;
    }

    /**
     * Gets the value of the returnResidentialZipCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResidentialZipCode() {
        return returnResidentialZipCode;
    }

    /**
     * Sets the value of the returnResidentialZipCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResidentialZipCode(Boolean value) {
        this.returnResidentialZipCode = value;
    }

    /**
     * Gets the value of the returnResourceClass property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceClass() {
        return returnResourceClass;
    }

    /**
     * Sets the value of the returnResourceClass property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceClass(Boolean value) {
        this.returnResourceClass = value;
    }

    /**
     * Gets the value of the returnResourceCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceCode() {
        return returnResourceCode;
    }

    /**
     * Sets the value of the returnResourceCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceCode(Boolean value) {
        this.returnResourceCode = value;
    }

    /**
     * Gets the value of the returnResourceDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceDesc() {
        return returnResourceDesc;
    }

    /**
     * Sets the value of the returnResourceDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceDesc(Boolean value) {
        this.returnResourceDesc = value;
    }

    /**
     * Gets the value of the returnResourceEffectiveDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnResourceEffectiveDate() {
        return returnResourceEffectiveDate;
    }

    /**
     * Sets the value of the returnResourceEffectiveDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnResourceEffectiveDate(Boolean value) {
        this.returnResourceEffectiveDate = value;
    }

    /**
     * Gets the value of the returnRetirementDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnRetirementDate() {
        return returnRetirementDate;
    }

    /**
     * Sets the value of the returnRetirementDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnRetirementDate(Boolean value) {
        this.returnRetirementDate = value;
    }

    /**
     * Gets the value of the returnSearchStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSearchStatus() {
        return returnSearchStatus;
    }

    /**
     * Sets the value of the returnSearchStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSearchStatus(Boolean value) {
        this.returnSearchStatus = value;
    }

    /**
     * Gets the value of the returnSeasonalInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSeasonalInd() {
        return returnSeasonalInd;
    }

    /**
     * Sets the value of the returnSeasonalInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSeasonalInd(Boolean value) {
        this.returnSeasonalInd = value;
    }

    /**
     * Gets the value of the returnSeasonalIndDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSeasonalIndDesc() {
        return returnSeasonalIndDesc;
    }

    /**
     * Sets the value of the returnSeasonalIndDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSeasonalIndDesc(Boolean value) {
        this.returnSeasonalIndDesc = value;
    }

    /**
     * Gets the value of the returnSecondName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSecondName() {
        return returnSecondName;
    }

    /**
     * Sets the value of the returnSecondName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSecondName(Boolean value) {
        this.returnSecondName = value;
    }

    /**
     * Gets the value of the returnServiceDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnServiceDate() {
        return returnServiceDate;
    }

    /**
     * Sets the value of the returnServiceDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnServiceDate(Boolean value) {
        this.returnServiceDate = value;
    }

    /**
     * Gets the value of the returnSigDateReason1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReason1() {
        return returnSigDateReason1;
    }

    /**
     * Sets the value of the returnSigDateReason1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReason1(Boolean value) {
        this.returnSigDateReason1 = value;
    }

    /**
     * Gets the value of the returnSigDateReason2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReason2() {
        return returnSigDateReason2;
    }

    /**
     * Sets the value of the returnSigDateReason2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReason2(Boolean value) {
        this.returnSigDateReason2 = value;
    }

    /**
     * Gets the value of the returnSigDateReason3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReason3() {
        return returnSigDateReason3;
    }

    /**
     * Sets the value of the returnSigDateReason3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReason3(Boolean value) {
        this.returnSigDateReason3 = value;
    }

    /**
     * Gets the value of the returnSigDateReason4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReason4() {
        return returnSigDateReason4;
    }

    /**
     * Sets the value of the returnSigDateReason4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReason4(Boolean value) {
        this.returnSigDateReason4 = value;
    }

    /**
     * Gets the value of the returnSigDateReasonDesc1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReasonDesc1() {
        return returnSigDateReasonDesc1;
    }

    /**
     * Sets the value of the returnSigDateReasonDesc1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReasonDesc1(Boolean value) {
        this.returnSigDateReasonDesc1 = value;
    }

    /**
     * Gets the value of the returnSigDateReasonDesc2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReasonDesc2() {
        return returnSigDateReasonDesc2;
    }

    /**
     * Sets the value of the returnSigDateReasonDesc2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReasonDesc2(Boolean value) {
        this.returnSigDateReasonDesc2 = value;
    }

    /**
     * Gets the value of the returnSigDateReasonDesc3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReasonDesc3() {
        return returnSigDateReasonDesc3;
    }

    /**
     * Sets the value of the returnSigDateReasonDesc3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReasonDesc3(Boolean value) {
        this.returnSigDateReasonDesc3 = value;
    }

    /**
     * Gets the value of the returnSigDateReasonDesc4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSigDateReasonDesc4() {
        return returnSigDateReasonDesc4;
    }

    /**
     * Sets the value of the returnSigDateReasonDesc4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSigDateReasonDesc4(Boolean value) {
        this.returnSigDateReasonDesc4 = value;
    }

    /**
     * Gets the value of the returnSignificantDate1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSignificantDate1() {
        return returnSignificantDate1;
    }

    /**
     * Sets the value of the returnSignificantDate1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSignificantDate1(Boolean value) {
        this.returnSignificantDate1 = value;
    }

    /**
     * Gets the value of the returnSignificantDate2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSignificantDate2() {
        return returnSignificantDate2;
    }

    /**
     * Sets the value of the returnSignificantDate2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSignificantDate2(Boolean value) {
        this.returnSignificantDate2 = value;
    }

    /**
     * Gets the value of the returnSignificantDate3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSignificantDate3() {
        return returnSignificantDate3;
    }

    /**
     * Sets the value of the returnSignificantDate3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSignificantDate3(Boolean value) {
        this.returnSignificantDate3 = value;
    }

    /**
     * Gets the value of the returnSignificantDate4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSignificantDate4() {
        return returnSignificantDate4;
    }

    /**
     * Sets the value of the returnSignificantDate4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSignificantDate4(Boolean value) {
        this.returnSignificantDate4 = value;
    }

    /**
     * Gets the value of the returnSiteCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSiteCode() {
        return returnSiteCode;
    }

    /**
     * Sets the value of the returnSiteCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSiteCode(Boolean value) {
        this.returnSiteCode = value;
    }

    /**
     * Gets the value of the returnSiteCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSiteCodeDesc() {
        return returnSiteCodeDesc;
    }

    /**
     * Sets the value of the returnSiteCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSiteCodeDesc(Boolean value) {
        this.returnSiteCodeDesc = value;
    }

    /**
     * Gets the value of the returnSkillsPassportCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportCode1() {
        return returnSkillsPassportCode1;
    }

    /**
     * Sets the value of the returnSkillsPassportCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportCode1(Boolean value) {
        this.returnSkillsPassportCode1 = value;
    }

    /**
     * Gets the value of the returnSkillsPassportCode2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportCode2() {
        return returnSkillsPassportCode2;
    }

    /**
     * Sets the value of the returnSkillsPassportCode2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportCode2(Boolean value) {
        this.returnSkillsPassportCode2 = value;
    }

    /**
     * Gets the value of the returnSkillsPassportCode3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportCode3() {
        return returnSkillsPassportCode3;
    }

    /**
     * Sets the value of the returnSkillsPassportCode3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportCode3(Boolean value) {
        this.returnSkillsPassportCode3 = value;
    }

    /**
     * Gets the value of the returnSkillsPassportType1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportType1() {
        return returnSkillsPassportType1;
    }

    /**
     * Sets the value of the returnSkillsPassportType1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportType1(Boolean value) {
        this.returnSkillsPassportType1 = value;
    }

    /**
     * Gets the value of the returnSkillsPassportType2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportType2() {
        return returnSkillsPassportType2;
    }

    /**
     * Sets the value of the returnSkillsPassportType2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportType2(Boolean value) {
        this.returnSkillsPassportType2 = value;
    }

    /**
     * Gets the value of the returnSkillsPassportType3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSkillsPassportType3() {
        return returnSkillsPassportType3;
    }

    /**
     * Sets the value of the returnSkillsPassportType3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSkillsPassportType3(Boolean value) {
        this.returnSkillsPassportType3 = value;
    }

    /**
     * Gets the value of the returnSmokerInd property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSmokerInd() {
        return returnSmokerInd;
    }

    /**
     * Sets the value of the returnSmokerInd property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSmokerInd(Boolean value) {
        this.returnSmokerInd = value;
    }

    /**
     * Gets the value of the returnSocialInsuranceNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialInsuranceNumber() {
        return returnSocialInsuranceNumber;
    }

    /**
     * Sets the value of the returnSocialInsuranceNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialInsuranceNumber(Boolean value) {
        this.returnSocialInsuranceNumber = value;
    }

    /**
     * Gets the value of the returnSocialSecurityNoType property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialSecurityNoType() {
        return returnSocialSecurityNoType;
    }

    /**
     * Sets the value of the returnSocialSecurityNoType property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialSecurityNoType(Boolean value) {
        this.returnSocialSecurityNoType = value;
    }

    /**
     * Gets the value of the returnSocialSecurityNoTypeDescription property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialSecurityNoTypeDescription() {
        return returnSocialSecurityNoTypeDescription;
    }

    /**
     * Sets the value of the returnSocialSecurityNoTypeDescription property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialSecurityNoTypeDescription(Boolean value) {
        this.returnSocialSecurityNoTypeDescription = value;
    }

    /**
     * Gets the value of the returnSocialSecurityNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSocialSecurityNumber() {
        return returnSocialSecurityNumber;
    }

    /**
     * Sets the value of the returnSocialSecurityNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSocialSecurityNumber(Boolean value) {
        this.returnSocialSecurityNumber = value;
    }

    /**
     * Gets the value of the returnStaffCategory property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnStaffCategory() {
        return returnStaffCategory;
    }

    /**
     * Sets the value of the returnStaffCategory property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnStaffCategory(Boolean value) {
        this.returnStaffCategory = value;
    }

    /**
     * Gets the value of the returnStaffCategoryDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnStaffCategoryDesc() {
        return returnStaffCategoryDesc;
    }

    /**
     * Sets the value of the returnStaffCategoryDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnStaffCategoryDesc(Boolean value) {
        this.returnStaffCategoryDesc = value;
    }

    /**
     * Gets the value of the returnStdTextRef property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnStdTextRef() {
        return returnStdTextRef;
    }

    /**
     * Sets the value of the returnStdTextRef property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnStdTextRef(Boolean value) {
        this.returnStdTextRef = value;
    }

    /**
     * Gets the value of the returnStdTextRefExists property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnStdTextRefExists() {
        return returnStdTextRefExists;
    }

    /**
     * Sets the value of the returnStdTextRefExists property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnStdTextRefExists(Boolean value) {
        this.returnStdTextRefExists = value;
    }

    /**
     * Gets the value of the returnSupplier property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSupplier() {
        return returnSupplier;
    }

    /**
     * Sets the value of the returnSupplier property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSupplier(Boolean value) {
        this.returnSupplier = value;
    }

    /**
     * Gets the value of the returnSupplierName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSupplierName() {
        return returnSupplierName;
    }

    /**
     * Sets the value of the returnSupplierName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSupplierName(Boolean value) {
        this.returnSupplierName = value;
    }

    /**
     * Gets the value of the returnSuspensionDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnSuspensionDate() {
        return returnSuspensionDate;
    }

    /**
     * Sets the value of the returnSuspensionDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnSuspensionDate(Boolean value) {
        this.returnSuspensionDate = value;
    }

    /**
     * Gets the value of the returnTerminationDate property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTerminationDate() {
        return returnTerminationDate;
    }

    /**
     * Sets the value of the returnTerminationDate property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTerminationDate(Boolean value) {
        this.returnTerminationDate = value;
    }

    /**
     * Gets the value of the returnTerminationReason property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTerminationReason() {
        return returnTerminationReason;
    }

    /**
     * Sets the value of the returnTerminationReason property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTerminationReason(Boolean value) {
        this.returnTerminationReason = value;
    }

    /**
     * Gets the value of the returnTerminationReasonDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTerminationReasonDesc() {
        return returnTerminationReasonDesc;
    }

    /**
     * Sets the value of the returnTerminationReasonDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTerminationReasonDesc(Boolean value) {
        this.returnTerminationReasonDesc = value;
    }

    /**
     * Gets the value of the returnThirdName property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnThirdName() {
        return returnThirdName;
    }

    /**
     * Sets the value of the returnThirdName property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnThirdName(Boolean value) {
        this.returnThirdName = value;
    }

    /**
     * Gets the value of the returnTitle property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTitle() {
        return returnTitle;
    }

    /**
     * Sets the value of the returnTitle property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTitle(Boolean value) {
        this.returnTitle = value;
    }

    /**
     * Gets the value of the returnTitleDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnTitleDesc() {
        return returnTitleDesc;
    }

    /**
     * Sets the value of the returnTitleDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnTitleDesc(Boolean value) {
        this.returnTitleDesc = value;
    }

    /**
     * Gets the value of the returnUnionCode property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUnionCode() {
        return returnUnionCode;
    }

    /**
     * Sets the value of the returnUnionCode property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUnionCode(Boolean value) {
        this.returnUnionCode = value;
    }

    /**
     * Gets the value of the returnUnionCodeDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUnionCodeDesc() {
        return returnUnionCodeDesc;
    }

    /**
     * Sets the value of the returnUnionCodeDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUnionCodeDesc(Boolean value) {
        this.returnUnionCodeDesc = value;
    }

    /**
     * Gets the value of the returnUserDefContact1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUserDefContact1() {
        return returnUserDefContact1;
    }

    /**
     * Sets the value of the returnUserDefContact1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUserDefContact1(Boolean value) {
        this.returnUserDefContact1 = value;
    }

    /**
     * Gets the value of the returnUserDefContact2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUserDefContact2() {
        return returnUserDefContact2;
    }

    /**
     * Sets the value of the returnUserDefContact2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUserDefContact2(Boolean value) {
        this.returnUserDefContact2 = value;
    }

    /**
     * Gets the value of the returnUserDefContact3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUserDefContact3() {
        return returnUserDefContact3;
    }

    /**
     * Sets the value of the returnUserDefContact3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUserDefContact3(Boolean value) {
        this.returnUserDefContact3 = value;
    }

    /**
     * Gets the value of the returnUserDefContact4 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUserDefContact4() {
        return returnUserDefContact4;
    }

    /**
     * Sets the value of the returnUserDefContact4 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUserDefContact4(Boolean value) {
        this.returnUserDefContact4 = value;
    }

    /**
     * Gets the value of the returnUserDefContact5 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnUserDefContact5() {
        return returnUserDefContact5;
    }

    /**
     * Sets the value of the returnUserDefContact5 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnUserDefContact5(Boolean value) {
        this.returnUserDefContact5 = value;
    }

    /**
     * Gets the value of the returnVeteranStatus property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVeteranStatus() {
        return returnVeteranStatus;
    }

    /**
     * Sets the value of the returnVeteranStatus property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVeteranStatus(Boolean value) {
        this.returnVeteranStatus = value;
    }

    /**
     * Gets the value of the returnVeteranStatusDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVeteranStatusDesc() {
        return returnVeteranStatusDesc;
    }

    /**
     * Sets the value of the returnVeteranStatusDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVeteranStatusDesc(Boolean value) {
        this.returnVeteranStatusDesc = value;
    }

    /**
     * Gets the value of the returnVisaCode1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode1() {
        return returnVisaCode1;
    }

    /**
     * Sets the value of the returnVisaCode1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode1(Boolean value) {
        this.returnVisaCode1 = value;
    }

    /**
     * Gets the value of the returnVisaCode1Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode1Desc() {
        return returnVisaCode1Desc;
    }

    /**
     * Sets the value of the returnVisaCode1Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode1Desc(Boolean value) {
        this.returnVisaCode1Desc = value;
    }

    /**
     * Gets the value of the returnVisaCode2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode2() {
        return returnVisaCode2;
    }

    /**
     * Sets the value of the returnVisaCode2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode2(Boolean value) {
        this.returnVisaCode2 = value;
    }

    /**
     * Gets the value of the returnVisaCode2Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode2Desc() {
        return returnVisaCode2Desc;
    }

    /**
     * Sets the value of the returnVisaCode2Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode2Desc(Boolean value) {
        this.returnVisaCode2Desc = value;
    }

    /**
     * Gets the value of the returnVisaCode3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode3() {
        return returnVisaCode3;
    }

    /**
     * Sets the value of the returnVisaCode3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode3(Boolean value) {
        this.returnVisaCode3 = value;
    }

    /**
     * Gets the value of the returnVisaCode3Desc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaCode3Desc() {
        return returnVisaCode3Desc;
    }

    /**
     * Sets the value of the returnVisaCode3Desc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaCode3Desc(Boolean value) {
        this.returnVisaCode3Desc = value;
    }

    /**
     * Gets the value of the returnVisaEffectiveDate1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaEffectiveDate1() {
        return returnVisaEffectiveDate1;
    }

    /**
     * Sets the value of the returnVisaEffectiveDate1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaEffectiveDate1(Boolean value) {
        this.returnVisaEffectiveDate1 = value;
    }

    /**
     * Gets the value of the returnVisaEffectiveDate2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaEffectiveDate2() {
        return returnVisaEffectiveDate2;
    }

    /**
     * Sets the value of the returnVisaEffectiveDate2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaEffectiveDate2(Boolean value) {
        this.returnVisaEffectiveDate2 = value;
    }

    /**
     * Gets the value of the returnVisaEffectiveDate3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaEffectiveDate3() {
        return returnVisaEffectiveDate3;
    }

    /**
     * Sets the value of the returnVisaEffectiveDate3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaEffectiveDate3(Boolean value) {
        this.returnVisaEffectiveDate3 = value;
    }

    /**
     * Gets the value of the returnVisaExpiryDate1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaExpiryDate1() {
        return returnVisaExpiryDate1;
    }

    /**
     * Sets the value of the returnVisaExpiryDate1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaExpiryDate1(Boolean value) {
        this.returnVisaExpiryDate1 = value;
    }

    /**
     * Gets the value of the returnVisaExpiryDate2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaExpiryDate2() {
        return returnVisaExpiryDate2;
    }

    /**
     * Sets the value of the returnVisaExpiryDate2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaExpiryDate2(Boolean value) {
        this.returnVisaExpiryDate2 = value;
    }

    /**
     * Gets the value of the returnVisaExpiryDate3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaExpiryDate3() {
        return returnVisaExpiryDate3;
    }

    /**
     * Sets the value of the returnVisaExpiryDate3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaExpiryDate3(Boolean value) {
        this.returnVisaExpiryDate3 = value;
    }

    /**
     * Gets the value of the returnVisaIssuedBy1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaIssuedBy1() {
        return returnVisaIssuedBy1;
    }

    /**
     * Sets the value of the returnVisaIssuedBy1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaIssuedBy1(Boolean value) {
        this.returnVisaIssuedBy1 = value;
    }

    /**
     * Gets the value of the returnVisaIssuedBy2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaIssuedBy2() {
        return returnVisaIssuedBy2;
    }

    /**
     * Sets the value of the returnVisaIssuedBy2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaIssuedBy2(Boolean value) {
        this.returnVisaIssuedBy2 = value;
    }

    /**
     * Gets the value of the returnVisaIssuedBy3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaIssuedBy3() {
        return returnVisaIssuedBy3;
    }

    /**
     * Sets the value of the returnVisaIssuedBy3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaIssuedBy3(Boolean value) {
        this.returnVisaIssuedBy3 = value;
    }

    /**
     * Gets the value of the returnVisaNumber1 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaNumber1() {
        return returnVisaNumber1;
    }

    /**
     * Sets the value of the returnVisaNumber1 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaNumber1(Boolean value) {
        this.returnVisaNumber1 = value;
    }

    /**
     * Gets the value of the returnVisaNumber2 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaNumber2() {
        return returnVisaNumber2;
    }

    /**
     * Sets the value of the returnVisaNumber2 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaNumber2(Boolean value) {
        this.returnVisaNumber2 = value;
    }

    /**
     * Gets the value of the returnVisaNumber3 property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnVisaNumber3() {
        return returnVisaNumber3;
    }

    /**
     * Sets the value of the returnVisaNumber3 property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnVisaNumber3(Boolean value) {
        this.returnVisaNumber3 = value;
    }

    /**
     * Gets the value of the returnWorkFacsimileNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkFacsimileNumber() {
        return returnWorkFacsimileNumber;
    }

    /**
     * Sets the value of the returnWorkFacsimileNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkFacsimileNumber(Boolean value) {
        this.returnWorkFacsimileNumber = value;
    }

    /**
     * Gets the value of the returnWorkGroup property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroup() {
        return returnWorkGroup;
    }

    /**
     * Sets the value of the returnWorkGroup property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroup(Boolean value) {
        this.returnWorkGroup = value;
    }

    /**
     * Gets the value of the returnWorkGroupCrew property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupCrew() {
        return returnWorkGroupCrew;
    }

    /**
     * Sets the value of the returnWorkGroupCrew property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupCrew(Boolean value) {
        this.returnWorkGroupCrew = value;
    }

    /**
     * Gets the value of the returnWorkGroupCrewDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupCrewDesc() {
        return returnWorkGroupCrewDesc;
    }

    /**
     * Sets the value of the returnWorkGroupCrewDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupCrewDesc(Boolean value) {
        this.returnWorkGroupCrewDesc = value;
    }

    /**
     * Gets the value of the returnWorkGroupDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkGroupDesc() {
        return returnWorkGroupDesc;
    }

    /**
     * Sets the value of the returnWorkGroupDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkGroupDesc(Boolean value) {
        this.returnWorkGroupDesc = value;
    }

    /**
     * Gets the value of the returnWorkLocation property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkLocation() {
        return returnWorkLocation;
    }

    /**
     * Sets the value of the returnWorkLocation property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkLocation(Boolean value) {
        this.returnWorkLocation = value;
    }

    /**
     * Gets the value of the returnWorkLocationDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkLocationDesc() {
        return returnWorkLocationDesc;
    }

    /**
     * Sets the value of the returnWorkLocationDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkLocationDesc(Boolean value) {
        this.returnWorkLocationDesc = value;
    }

    /**
     * Gets the value of the returnWorkMobilePhoneNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkMobilePhoneNumber() {
        return returnWorkMobilePhoneNumber;
    }

    /**
     * Sets the value of the returnWorkMobilePhoneNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkMobilePhoneNumber(Boolean value) {
        this.returnWorkMobilePhoneNumber = value;
    }

    /**
     * Gets the value of the returnWorkOrderPrefix property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkOrderPrefix() {
        return returnWorkOrderPrefix;
    }

    /**
     * Sets the value of the returnWorkOrderPrefix property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkOrderPrefix(Boolean value) {
        this.returnWorkOrderPrefix = value;
    }

    /**
     * Gets the value of the returnWorkOrderPrefixDesc property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkOrderPrefixDesc() {
        return returnWorkOrderPrefixDesc;
    }

    /**
     * Sets the value of the returnWorkOrderPrefixDesc property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkOrderPrefixDesc(Boolean value) {
        this.returnWorkOrderPrefixDesc = value;
    }

    /**
     * Gets the value of the returnWorkTelephoneExtension property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkTelephoneExtension() {
        return returnWorkTelephoneExtension;
    }

    /**
     * Sets the value of the returnWorkTelephoneExtension property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkTelephoneExtension(Boolean value) {
        this.returnWorkTelephoneExtension = value;
    }

    /**
     * Gets the value of the returnWorkTelephoneNumber property.
     * 
     * @return
     *     possible object is
     *     {@link Boolean }
     *     
     */
    public Boolean isReturnWorkTelephoneNumber() {
        return returnWorkTelephoneNumber;
    }

    /**
     * Sets the value of the returnWorkTelephoneNumber property.
     * 
     * @param value
     *     allowed object is
     *     {@link Boolean }
     *     
     */
    public void setReturnWorkTelephoneNumber(Boolean value) {
        this.returnWorkTelephoneNumber = value;
    }

}
