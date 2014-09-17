# This program takes Excel file(s) containing information about various resources and
# writes them to XML for import into geoportal

import sys
from owslib.etree import etree
import xlrd  # xlrd to extract values from cells
import uuid  # guid used for file ID

SubElement = etree.SubElement
Element = etree.Element
toString = etree.tostring

# this block is used to reference code dictionaries
#e = etree.fromstring(urlopen('http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml').read())
#c = CodelistCatalogue(e)
#c.getcodelistdictionaries()
#c.getcodedefinitionidentifiers('MD_SpatialRepresentationTypeCode')
#c.getcodedefinitionidentifiers('MD_TopicCategoryCode')

#open workbook, reference first and only sheet
wb = xlrd.open_workbook("EarthCube Resources.xlsx", sys.stdout)
sh = wb.sheet_by_index(0)


# Removes Graphic Overview element if resources has no explicit graphic attached to it
def remove_graphic(md_data_id, graphic_overview):
    if fileNameString.text == "":
        md_data_id.remove(graphic_overview)


# Removes thesaurus element if resource has no thesaurus
def remove_thesaurus(md_data_id, descriptivekeywords):
    if thesaurusTitle.text == "":
        MD_Keywords.remove(thesaurusName)


# Removes constraint element and all its subelements if resource has no constraints
def remove_constraints(md_data_id, resource_constraints):
    if useLimitationMDString.text == "":
        resource_constraints.remove(MD_Constraints)
    if useLimitString.text == "":
        MD_LegalConstraints.remove(useLimit)
    if access_RestrictionCode.text == "":
        MD_LegalConstraints.remove(accessConstraints)
    if use_RestrictionCode.text == "":
        MD_LegalConstraints.remove(useConstraints)
    if otherString.text == "":
        MD_LegalConstraints.remove(otherConstraints)
    if useLimitString.text == "" and access_RestrictionCode.text == "" and use_RestrictionCode.text == "" \
            and otherString.text == "":
        resource_constraints.remove(MD_LegalConstraints)
    if useLimitationString.text == "":
        MD_SecurityConstraints.remove(useLimitation)
    if MD_ClassificationCode.text == "":
        MD_SecurityConstraints.remove(classification)
    if userNoteString.text == "":
        MD_SecurityConstraints.remove(userNote)
    if classSysString.text == "":
        MD_SecurityConstraints.remove(classificationSystem)
    if handlingString.text == "":
        MD_SecurityConstraints.remove(handlingDescription)
    if useLimitationString.text == "" and MD_ClassificationCode.text == "" and userNoteString.text == "" \
            and classSysString.text == "" and handlingString.text == "":
        resource_constraints.remove(MD_SecurityConstraints)
    return md_data_id


# Removes data quality data if resources has no explicit report
# Also necessary if resource is actually a database to other data sets
def remove_quality(tree_root):
    if MD_ScopeCode.text == "":
        tree_root.remove(dataQualityInfo)
    return tree_root


# Using comma as a delimiter, separates each keyword into its own subelement
def separate_keywords(md_keywords, keywords_entered, length):
    for y in range(0, length):
        keyword = SubElement(md_keywords, '{http://www.isotc211.org/2005/gmd}keyword')
        keyword_string = SubElement(keyword, '{http://www.isotc211.org/2005/gco}CharacterString')
        keyword_string.text = keywords_entered[y]
        md_keywords.append(keyword)
    return md_keywords


def make_topic_categories(parent_elem, categories_entered, length):
    for x in range(0, length):
        topic_category = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}topicCategory')
        md_topic_category_code = SubElement(topic_category, '{http://www.isotc211.org/2005/gmd}MD_TopicCategoryCode')
        a_category = categories_entered[x]
        md_topic_category_code.text = str(match_topic_category(a_category))
        parent_elem.append(topic_category)
    return parent_elem


def match_topic_category(a_category):
    return {
        'Geological and Geophysical': 'geoscientificInformation',
        'Agriculture and Farming': 'farming',
        'Elevation and Derived Products': 'elevation',
        'Utilities and Communication': 'utilitiesCommunication',
        'Oceans and Estuaries': 'oceans',
        'Administrative and Political Boundaries': 'boundaries',
        'Inland Water Resources': 'inlandWaters',
        'Military': 'intelligenceMilitary',
        'Environment and Conservation': 'environment',
        'Locations and Geodetic Networks': 'location',
        'Business and Economic': 'economy',
        'Cadastral': 'planningCadastre',
        'Biology and Ecology': 'biota',
        'Human Health and Disease': 'health',
        'Imagery and Base Maps': 'imageryBaseMapsEarthCover',
        'Transportation Networks': 'transportation',
        'Cultural, Society and Demography': 'society',
        'Facilities and Structures': 'structure',
        'Atmosphere and Climatic': 'climatologyMeteorologyAtmosphere'
    }[a_category]

start_row = 1
end_row = 3
for i in range(start_row, end_row):  # range of resources; starts at 1 because first row is column headers
    #try:
    #    etree.parse("resource_%i.xml" % i)
    #    print("File found for resource_%i. File identifier will be retained" % i)
    #    root = etree.parse("resource_%i.xml" % i).getroot()
    #    fileExists = True
    #except IOError:
    #    print("File not found for resource_%i.xml. Creating new file..." % i)
    #    fileExists = False
    root = Element('{http://www.isotc211.org/2005/gmd}MD_Metadata', {'gmd': 'http://www.isotc211.org/2005/gmd',
                                                                         'gco': 'http://www.isotc211.org/2005/gco',
                                                                         'gml': 'http://www.opengis.net/gml/3.2',
                                                                         'srv': 'http://www.isotc211.org/2005/srv'})
    fileIdentifier = SubElement(root, '{http://www.isotc211.org/2005/gmd}fileIdentifier')
    fileID = SubElement(fileIdentifier, '{http://www.isotc211.org/2005/gco}CharacterString')
    #if fileExists:
    #    fileID.text = root.findtext('{http://www.isotc211.org/2005/gmd}fileIdentifier/'
    #                                '{http://www.isotc211.org/2005/gco}CharacterString')
    #    print("File Identifier: " + fileID.text)
    #else:
    fileID.text = str(uuid.uuid4())
    language = SubElement(root, '{http://www.isotc211.org/2005/gmd}language')
    languageString = SubElement(language, '{http://www.isotc211.org/2005/gco}CharacterString')
    languageString.text = sh.cell_value(i, 3)
    hierarchyLevel = SubElement(root, '{http://www.isotc211.org/2005/gmd}hierarchyLevel')
    # Code elements with restricted values extract their appropriate value from spreadsheet.
    # Element text is then assigned to codeListValue
    MD_ScopeCode = SubElement(hierarchyLevel, '{http://www.isotc211.org/2005/gmd}MD_ScopeCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_ScopeCode',
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 4)})
    MD_ScopeCode.text = MD_ScopeCode.get('codeListValue')
    contact = SubElement(root, '{http://www.isotc211.org/2005/gmd}contact')
    CI_ResponsibleParty = SubElement(contact, '{http://www.isotc211.org/2005/gmd}CI_ResponsibleParty')
    organisationName = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}organisationName')
    orgName = SubElement(organisationName, '{http://www.isotc211.org/2005/gco}CharacterString')
    orgName.text = sh.cell_value(i, 5)
    contactInfo = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}contactInfo')
    CI_Contact = SubElement(contactInfo, '{http://www.isotc211.org/2005/gmd}CI_Contact')
    address = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}address')
    CI_Address = SubElement(address, '{http://www.isotc211.org/2005/gmd}CI_Address')
    electronicMailAddress = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}electronicMailAddress')
    elecMailAddress = SubElement(electronicMailAddress, '{http://www.isotc211.org/2005/gco}CharacterString')
    elecMailAddress.text = sh.cell_value(i, 6)
    role = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}role')
    CI_RoleCode = SubElement(role, '{http://www.isotc211.org/2005/gmd}CI_RoleCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_RoleCode',
        'codeSpace': 'ISOTC211/19115', 'codeListValue': sh.cell_value(i, 7)})
    CI_RoleCode.text = CI_RoleCode.get('codeListValue')
    dateStamp = SubElement(root, '{http://www.isotc211.org/2005/gmd}dateStamp')
    metadataDate = SubElement(dateStamp, '{http://www.isotc211.org/2005/gco}Date')
    # When using xlrd, dates are returned as doubles.
    # To get the proper date, use xlrd's date to tuple function: xldate_as_tuple
    metadataDateTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 8), wb.datemode))  # extract date from cell,
    # convert to Gregorian time
    metadataDateList = list(map(int, metadataDateTuple))  # turn xlrd date into a list of three integers
                                                          # each element is either year, month, or day
    metadataYear = str(metadataDateList[0])  # make year a string
    if metadataDateList[1] < 10:  # check if month is less than 10, if so, affix 0 to beginning
        metadataMonth = '0' + str(metadataDateList[1])
    else:
        metadataMonth = str(metadataDateList[1])
    if metadataDateList[2] < 10:
        metadataDay = '0' + str(metadataDateList[2])
    else:
        metadataDay = str(metadataDateList[2])
    metadataDate.text = metadataYear + '-' + metadataMonth + '-' + metadataDay
    metadataStandardName = SubElement(root, '{http://www.isotc211.org/2005/gmd}metadataStandardName')
    stdName = SubElement(metadataStandardName, '{http://www.isotc211.org/2005/gco}CharacterString')
    stdName.text = sh.cell_value(i, 9)
    metadataStandardVersion = SubElement(root, '{http://www.isotc211.org/2005/gmd}metadataStandardVersion')
    stdVer = SubElement(metadataStandardVersion, '{http://www.isotc211.org/2005/gco}CharacterString')
    stdVer.text = str(sh.cell_value(i, 10)).strip('.0')
    dataSetURI = SubElement(root, '{http://www.isotc211.org/2005/gmd}dataSetURI')
    dataSetURIString = SubElement(dataSetURI, '{http://www.isotc211.org/2005/gco}CharacterString')
    dataSetURIString.text = sh.cell_value(i, 68)
    locale = SubElement(root, '{http://www.isotc211.org/2005/gmd}locale')
    PT_Locale = SubElement(locale, '{http://www.isotc211.org/2005/gmd}PT_Locale')
    languageCode = SubElement(PT_Locale, '{http://www.isotc211.org/2005/gmd}languageCode')
    LanguageCode = SubElement(languageCode, '{http://www.isotc211.org/2005/gmd}LanguageCode',
                              {'codeListValue': 'en', 'codeList': '#LanguageCode'})
    lang_country = SubElement(PT_Locale, '{http://www.isotc211.org/2005/gmd}country')
    characterEncoding = SubElement(PT_Locale, '{http://www.isotc211.org/2005/gmd}characterEncoding')
    MD_CharacterSetCode = SubElement(characterEncoding, '{http://www.isotc211.org/2005/gmd}MD_CharacterSetCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_CharacterSetCode',
        'codeListValue': "utf8"})
    spatialRepresentationInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}spatialRepresentationInfo')
    referenceSystemInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}referenceSystemInfo')
    MD_ReferenceSystem = SubElement(referenceSystemInfo, '{http://www.isotc211.org/2005/gmd}MD_ReferenceSystem')
    referenceSystemIdentifier = SubElement(MD_ReferenceSystem,
                                           '{http://www.isotc211.org/2005/gmd}referenceSystemIdentifier')
    RS_Identifier = SubElement(referenceSystemIdentifier, '{http://www.isotc211.org/2005/gmd}RS_Identifier')
    code = SubElement(RS_Identifier, '{http://www.isotc211.org/2005/gmd}code')
    codeString = SubElement(code, '{http://www.isotc211.org/2005/gco}CharacterString')
    codeString.text = "WGS84"
    codeSpace = SubElement(RS_Identifier, '{http://www.isotc211.org/2005/gmd}codeSpace')
    codeSpaceString = SubElement(codeSpace, '{http://www.isotc211.org/2005/gco}CharacterString')
    codeSpaceString.text = "OGP Geomatics Committee"
    metadataExtensionInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}metadataExtensionInfo')
    identificationInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}identificationInfo')
    MD_dataID = SubElement(identificationInfo, '{http://www.isotc211.org/2005/gmd}MD_DataIdentification')
    citation = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}citation')
    CI_Citation = SubElement(citation, '{http://www.isotc211.org/2005/gmd}CI_Citation')
    title = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}title')
    resourceTitle = SubElement(title, '{http://www.isotc211.org/2005/gco}CharacterString')
    resourceTitle.text = sh.cell_value(i, 12)
    date1 = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}date')
    CI_Date = SubElement(date1, '{http://www.isotc211.org/2005/gmd}CI_Date')
    date2 = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}date')
    resourceDate = SubElement(date2, '{http://www.isotc211.org/2005/gco}Date')
    #Using 2013 as a marker since this metadata was made in 2013. If only the year is know for a resource's date,
    #it will logically have to be less than 2013. The year is grabbed and the .0 that results from xlrd is stripped
    if sh.cell_value(i, 13) < 2013:
        resourceDate.text = str(sh.cell_value(i, 13)).strip('.0')
    else:
        resourceDateTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 13), wb.datemode))
        resourceDateList = list(map(int, resourceDateTuple))
        resourceYear = str(resourceDateList[0])
        if resourceDateList[1] < 10:
            resourceMonth = '0' + str(resourceDateList[1])
        else:
            resourceMonth = str(resourceDateList[1])
        if resourceDateList[2] < 10:
            resourceDay = '0' + str(resourceDateList[2])
        else:
            resourceDay = str(resourceDateList[2])
        resourceDate.text = resourceYear + '-' + resourceMonth + '-' + resourceDay
    resourceDateType = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}dateType')
    CI_DateTypeCode = SubElement(resourceDateType, '{http://www.isotc211.org/2005/gmd}CI_DateTypeCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_DateTypeCode',
        'codeSpace': 'ISOTC211/19115', 'codeListValue': sh.cell_value(i, 14)})
    CI_DateTypeCode.text = CI_DateTypeCode.get('codeListValue')
    identifier = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}identifier')
    RS_Identifier = SubElement(identifier, '{http://www.isotc211.org/2005/gmd}RS_Identifier')
    URI_code = SubElement(RS_Identifier, '{http://www.isotc211.org/2005/gmd}code')
    URIString = SubElement(URI_code, '{http://www.isotc211.org/2005/gco}CharacterString')
    URIString.text = sh.cell_value(i, 15)
    codeSpace = SubElement(RS_Identifier, '{http://www.isotc211.org/2005/gmd}codeSpace')
    codeSpaceString = SubElement(codeSpace, '{http://www.isotc211.org/2005/gco}CharacterString')
    codeSpaceString.text = sh.cell_value(i, 16)
    abstract = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}abstract')
    Abstract = SubElement(abstract, '{http://www.isotc211.org/2005/gco}CharacterString')
    Abstract.text = sh.cell_value(i, 17)
    pointOfContact = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}pointOfContact')
    CI_ResponsibleParty = SubElement(pointOfContact, '{http://www.isotc211.org/2005/gmd}CI_ResponsibleParty')
    individualName = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}individualName')
    indivName = SubElement(individualName, '{http://www.isotc211.org/2005/gco}CharacterString')
    indivName.text = sh.cell_value(i, 18)
    organisationName = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}organisationName')
    orgName = SubElement(organisationName, '{http://www.isotc211.org/2005/gco}CharacterString')
    orgName.text = sh.cell_value(i, 19)
    positionName = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}positionName')
    posName = SubElement(positionName, '{http://www.isotc211.org/2005/gco}CharacterString')
    posName.text = sh.cell_value(i, 20)
    contactInfo = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}contactInfo')
    CI_Contact = SubElement(contactInfo, '{http://www.isotc211.org/2005/gmd}CI_Contact')
    phone = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}phone')
    CI_Telephone = SubElement(phone, '{http://www.isotc211.org/2005/gmd}CI_Telephone')
    voice = SubElement(CI_Telephone, '{http://www.isotc211.org/2005/gmd}voice')
    voicePhone = SubElement(voice, '{http://www.isotc211.org/2005/gco}CharacterString')
    voicePhone.text = sh.cell_value(i, 21)
    facsimile = SubElement(CI_Telephone, '{http://www.isotc211.org/2005/gmd}facsimile')
    fax = SubElement(facsimile, '{http://www.isotc211.org/2005/gco}CharacterString')
    fax_cell = sh.cell_value(i, 22)
    address = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}address')
    CI_Address = SubElement(address, '{http://www.isotc211.org/2005/gmd}CI_Address')
    deliveryPoint = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}deliveryPoint')
    delivery = SubElement(deliveryPoint, '{http://www.isotc211.org/2005/gco}CharacterString')
    delivery.text = sh.cell_value(i, 23)
    city = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}city')
    City = SubElement(city, '{http://www.isotc211.org/2005/gco}CharacterString')
    City.text = sh.cell_value(i, 24)
    administrativeArea = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}administrativeArea')
    adminArea = SubElement(administrativeArea, '{http://www.isotc211.org/2005/gco}CharacterString')
    adminArea.text = str(sh.cell_value(i, 25))
    postalCode = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}postalCode')
    postCode = SubElement(postalCode, '{http://www.isotc211.org/2005/gco}CharacterString')
    postCode.text = str(sh.cell_value(i, 26))
    country = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}country')
    nation = SubElement(country, '{http://www.isotc211.org/2005/gco}CharacterString')
    nation.text = sh.cell_value(i, 27)
    eMailAddress = SubElement(CI_Address, '{http://www.isotc211.org/2005/gmd}electronicMailAddress')
    eMailAddress = SubElement(eMailAddress, '{http://www.isotc211.org/2005/gco}CharacterString')
    eMailAddress.text = sh.cell_value(i, 28)
    onlineResource = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}onlineResource')
    CI_OnlineResource = SubElement(onlineResource, '{http://www.isotc211.org/2005/gmd}CI_OnlineResource')
    linkage = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}linkage')
    URL = SubElement(linkage, '{http://www.isotc211.org/2005/gmd}URL')
    URL.text = sh.cell_value(i, 29)
    protocol = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}protocol')
    protocolString = SubElement(protocol, '{http://www.isotc211.org/2005/gco}CharacterString')
    protocolString.text = sh.cell_value(i, 30)
    applicationProfile = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}applicationProfile')
    appProfile = SubElement(applicationProfile, '{http://www.isotc211.org/2005/gco}CharacterString')
    appProfile.text = sh.cell_value(i, 31)
    name = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}name')
    nameString = SubElement(name, '{http://www.isotc211.org/2005/gco}CharacterString')
    nameString.text = sh.cell_value(i, 32)
    description = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}description')
    descriptionString = SubElement(description, '{http://www.isotc211.org/2005/gco}CharacterString')
    descriptionString.text = sh.cell_value(i, 33)
    function = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}function')
    CI_OnlineFunctionCode = SubElement(function, '{http://www.isotc211.org/2005/gmd}CI_OnLineFunctionCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_OnLineFunctionCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 34)})
    CI_OnlineFunctionCode.text = CI_OnlineFunctionCode.get('codeListValue')
    hoursOfService = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}hoursOfService')
    hours = SubElement(hoursOfService, '{http://www.isotc211.org/2005/gco}CharacterString')
    hours.text = sh.cell_value(i, 35)
    contactInstructions = SubElement(CI_Contact, '{http://www.isotc211.org/2005/gmd}contactInstructions')
    contactInstr = SubElement(contactInstructions, '{http://www.isotc211.org/2005/gco}CharacterString')
    contactInstr.text = sh.cell_value(i, 36)
    role = SubElement(CI_ResponsibleParty, '{http://www.isotc211.org/2005/gmd}role')
    CI_RoleCode = SubElement(role, '{http://www.isotc211.org/2005/gmd}CI_RoleCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_RoleCode',
        'codeSpace': 'ISOTC211/19115', 'codeListValue': sh.cell_value(i, 37)})
    CI_RoleCode.text = CI_RoleCode.get('codeListValue')
    graphicOverview = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}graphicOverview')
    MD_BrowseGraphic = SubElement(graphicOverview, '{http://www.isotc211.org/2005/gmd}MD_BrowseGraphic')
    fileName = SubElement(MD_BrowseGraphic, '{http://www.isotc211.org/2005/gmd}fileName')
    fileNameString = SubElement(fileName, '{http://www.isotc211.org/2005/gco}CharacterString')
    fileNameString.text = sh.cell_value(i, 38)
    fileDescription_0 = SubElement(MD_BrowseGraphic, '{http://www.isotc211.org/2005/gmd}fileDescription')
    fileDescription = SubElement(fileDescription_0, '{http://www.isotc211.org/2005/gco}CharacterString')
    fileDescriptionString = sh.cell_value(i, 39)
    fileType = SubElement(MD_BrowseGraphic, '{http://www.isotc211.org/2005/gmd}fileType')
    fileTypeString = SubElement(fileType, '{http://www.isotc211.org/2005/gco}CharacterString')
    fileTypeString.text = sh.cell_value(i, 40)
    descriptiveKeywords = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}descriptiveKeywords')
    MD_Keywords = SubElement(descriptiveKeywords, '{http://www.isotc211.org/2005/gmd}MD_Keywords')
    keywordsEntered = sh.cell_value(i, 41).split('; ')
    numberKeywordsEntered = len(keywordsEntered)
    separate_keywords(MD_Keywords, keywordsEntered, numberKeywordsEntered)
    keywordType = SubElement(MD_Keywords, '{http://www.isotc211.org/2005/gmd}type')
    MD_KeywordTypeCode = SubElement(keywordType, '{http://www.isotc211.org/2005/gmd}MD_KeywordTypeCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_KeywordTypeCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 42)})
    MD_KeywordTypeCode.text = MD_KeywordTypeCode.get('codeListValue')
    thesaurusName = SubElement(MD_Keywords, '{http://www.isotc211.org/2005/gmd}thesaurusName')
    CI_Citation = SubElement(thesaurusName, '{http://www.isotc211.org/2005/gmd}CI_Citation')
    title = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}title')
    thesaurusTitle = SubElement(title, '{http://www.isotc211.org/2005/gco}CharacterString')
    thesaurusTitle.text = sh.cell_value(i, 43)
    dateForDatesSake = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}date')
    CI_Date = SubElement(dateForDatesSake, '{http://www.isotc211.org/2005/gmd}CI_Date')
    realDate = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}date')
    thesaurusDate = SubElement(realDate, '{http://www.isotc211.org/2005/gco}Date')
    if sh.cell_value(i, 44) == "":
        thesaurusDate.text = ""
    else:
        if sh.cell_value(i, 44) < 2013:
            thesaurusDate.text = str(sh.cell_value(i, 44)).strip('.0')
        else:
            thesaurusDateTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 44), wb.datemode))
            thesaurusDateList = list(map(int, thesaurusDateTuple))
            thesaurusYear = str(thesaurusDateList[0])
            if thesaurusDateList[1] < 10:
                thesaurusMonth = '0' + str(thesaurusDateList[1])
            else:
                thesaurusMonth = str(thesaurusDateList[1])
            if thesaurusDateList[2] < 10:
                thesaurusDay = '0' + str(thesaurusDateList[2])
            else:
                thesaurusDay = str(thesaurusDateList[2])
            thesaurusDate.text = thesaurusYear + '-' + thesaurusMonth + '-' + thesaurusDay
    dateType = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}dateType')
    CI_DateTypeCode = SubElement(dateType, '{http://www.isotc211.org/2005/gmd}CI_DateTypeCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_DateTypeCode',
        'codeSpace': 'ISOTC211/19115', 'codeListValue': sh.cell_value(i, 45)})
    CI_DateTypeCode.text = CI_DateTypeCode.get('codeListValue')
    resourceConstraints = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}resourceConstraints')
    MD_Constraints = SubElement(resourceConstraints, '{http://www.isotc211.org/2005/gmd}MD_Constraints')
    useLimitationMD = SubElement(MD_Constraints, '{http://www.isotc211.org/2005/gmd}useLimitation')
    useLimitationMDString = SubElement(useLimitationMD, '{http://www.isotc211.org/2005/gco}CharacterString')
    useLimitationMDString.text = sh.cell_value(i, 46)
    MD_LegalConstraints = SubElement(resourceConstraints, '{http://www.isotc211.org/2005/gmd}MD_LegalConstraints')
    useLimit = SubElement(MD_LegalConstraints, '{http://www.isotc211.org/2005/gmd}useLimitation')
    useLimitString = SubElement(useLimit, '{http://www.isotc211.org/2005/gco}CharacterString')
    useLimitString.text = sh.cell_value(i, 47)
    accessConstraints = SubElement(MD_LegalConstraints, '{http://www.isotc211.org/2005/gmd}accessConstraints')
    access_RestrictionCode = SubElement(accessConstraints, '{http://www.isotc211.org/2005/gmd}MD_RestrictionCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_RestrictionCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 48)})
    access_RestrictionCode.text = access_RestrictionCode.get('codeListValue')
    useConstraints = SubElement(MD_LegalConstraints, '{http://www.isotc211.org/2005/gmd}useConstraints')
    use_RestrictionCode = SubElement(useConstraints, '{http://www.isotc211.org/2005/gmd}MD_RestrictionCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_RestrictionCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 49)})
    use_RestrictionCode.text = use_RestrictionCode.get('codeListValue')
    otherConstraints = SubElement(MD_LegalConstraints, '{http://www.isotc211.org/2005/gmd}otherConstraints')
    otherString = SubElement(otherConstraints, '{http://www.isotc211.org/2005/gco}CharacterString')
    otherString.text = sh.cell_value(i, 50)
    MD_SecurityConstraints = SubElement(resourceConstraints, '{http://www.isotc211.org/2005/gmd}MD_SecurityConstraints')
    useLimitation = SubElement(MD_SecurityConstraints, '{http://www.isotc211.org/2005/gmd}useLimitation')
    useLimitationString = SubElement(useLimitation, '{http://www.isotc211.org/2005/gco}CharacterString')
    useLimitationString.text = sh.cell_value(i, 51)
    classification = SubElement(MD_SecurityConstraints, '{http://www.isotc211.org/2005/gmd}classification')
    MD_ClassificationCode = SubElement(classification, '{http://www.isotc211.org/2005/gmd}MD_ClassificationCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_ClassificationCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 52)})
    MD_ClassificationCode.text = MD_ClassificationCode.get('codeListValue')
    userNote = SubElement(MD_SecurityConstraints, '{http://www.isotc211.org/2005/gmd}userNote')
    userNoteString = SubElement(userNote, '{http://www.isotc211.org/2005/gco}CharacterString')
    userNoteString.text = sh.cell_value(i, 53)
    classificationSystem = SubElement(MD_SecurityConstraints, '{http://www.isotc211.org/2005/gmd}classificationSystem')
    classSysString = SubElement(classificationSystem, '{http://www.isotc211.org/2005/gco}CharacterString')
    classSysString.text = sh.cell_value(i, 54)
    handlingDescription = SubElement(MD_SecurityConstraints, '{http://www.isotc211.org/2005/gmd}handlingDescription')
    handlingString = SubElement(handlingDescription, '{http://www.isotc211.org/2005/gco}CharacterString')
    handlingString.text = sh.cell_value(i, 55)
    spatialRepType = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}spatialRepresentationType')
    spatialRepTypeCode = SubElement(spatialRepType,
                                    '{http://www.isotc211.org/2005/gmd}MD_SpatialRepresentationTypeCode',
                                    {'codeList':
                                    "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#"
                                    "MD_SpatialRepresentationTypeCode",
                                    'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 56)})
    spatialRepTypeCode.text = spatialRepTypeCode.get('codeListValue')
    spatialResolution = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}spatialResolution')
    MD_Resolution = SubElement(spatialResolution, '{http://www.isotc211.org/2005/gmd}MD_Resolution')
    equivalentScale = SubElement(MD_Resolution, '{http://www.isotc211.org/2005/gmd}equivalentScale')
    MD_RepresentativeFraction = SubElement(equivalentScale,
                                           '{http://www.isotc211.org/2005/gmd}MD_RepresentativeFraction')
    denominator = SubElement(MD_RepresentativeFraction, '{http://www.isotc211.org/2005/gmd}denominator')
    denominatorInt = SubElement(denominator, '{http://www.isotc211.org/2005/gco}Integer')
    if sh.cell_value(i, 57) is not "":
        denominatorInt.text = str(int(sh.cell_value(i, 57)))
    elif sh.cell_value(i, 58) is not "":
        MD_Resolution.remove(equivalentScale)
        distance = SubElement(MD_Resolution, '{http://www.isotc211.org/2005/gmd}distance')
        Distance = SubElement(distance, '{http://www.isotc211.org/2005/gco}Distance', {'uom': sh.cell_value(i, 59)})
        Distance.text = str(int(sh.cell_value(i, 58)))
    else:
        MD_dataID.remove(spatialResolution)
    language = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}language')
    languageString = SubElement(language, '{http://www.isotc211.org/2005/gco}CharacterString')
    languageString.text = sh.cell_value(i, 60)
    categoriesEntered = sh.cell_value(i, 61).split('; ')
    numberEntered = len(categoriesEntered)
    make_topic_categories(MD_dataID, categoriesEntered, numberEntered)
    extent = SubElement(MD_dataID, '{http://www.isotc211.org/2005/gmd}extent')
    EX_Extent = SubElement(extent, '{http://www.isotc211.org/2005/gmd}EX_Extent')
    geographicElement = SubElement(EX_Extent, '{http://www.isotc211.org/2005/gmd}geographicElement')
    EX_GeographicBoundBox = SubElement(geographicElement, '{http://www.isotc211.org/2005/gmd}EX_GeographicBoundingBox')
    westBoundingLongitude = SubElement(EX_GeographicBoundBox, '{http://www.isotc211.org/2005/gmd}westBoundLongitude')
    Decimal = SubElement(westBoundingLongitude, '{http://www.isotc211.org/2005/gco}Decimal')
    Decimal.text = str(sh.cell_value(i, 62))
    eastBoundingLongitude = SubElement(EX_GeographicBoundBox, '{http://www.isotc211.org/2005/gmd}eastBoundLongitude')
    Decimal = SubElement(eastBoundingLongitude, '{http://www.isotc211.org/2005/gco}Decimal')
    Decimal.text = str(sh.cell_value(i, 63))
    southBoundingLatitude = SubElement(EX_GeographicBoundBox, '{http://www.isotc211.org/2005/gmd}southBoundLatitude')
    Decimal = SubElement(southBoundingLatitude, '{http://www.isotc211.org/2005/gco}Decimal')
    Decimal.text = str(sh.cell_value(i, 64))
    northBoundingLatitude = SubElement(EX_GeographicBoundBox, '{http://www.isotc211.org/2005/gmd}northBoundLatitude')
    Decimal = SubElement(northBoundingLatitude, '{http://www.isotc211.org/2005/gco}Decimal')
    Decimal.text = str(sh.cell_value(i, 65))
    temporalElement = SubElement(EX_Extent, '{http://www.isotc211.org/2005/gmd}temporalElement')
    EX_TemporalExtent = SubElement(temporalElement, '{http://www.isotc211.org/2005/gmd}EX_TemporalExtent')
    extent = SubElement(EX_TemporalExtent, '{http://www.isotc211.org/2005/gmd}extent')
    TimePeriod = SubElement(extent, '{http://www.opengis.net/gml/3.2}TimePeriod',
                            {'{http://www.opengis.net/gml/3.2}id': "Temporal"})
    beginPosition = SubElement(TimePeriod, '{http://www.opengis.net/gml/3.2}beginPosition')
    if sh.cell_value(i, 66) == "":
        beginPosition.text = ""
    else:
        if sh.cell_value(i, 66) < 2013:
            beginPosition.text = str(sh.cell_value(i, 66)).strip('.0')
        else:
            beginPositionTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 66), wb.datemode))
            beginPositionList = list(map(int, beginPositionTuple))
            beginPositionYear = str(beginPositionList[0])
            if beginPositionList[1] < 10:
                beginPositionMonth = '0' + str(beginPositionList[1])
            else:
                beginPositionMonth = str(beginPositionList[1])
            if beginPositionList[2] < 10:
                beginPositionDay = '0' + str(beginPositionList[2])
            else:
                beginPositionDay = str(beginPositionList[2])
            beginPosition.text = beginPositionYear + '-' + beginPositionMonth + '-' + beginPositionDay
    endPosition = SubElement(TimePeriod, '{http://www.opengis.net/gml/3.2}endPosition')
    if sh.cell_value(i, 67) == "":
        endPosition.text = ""
    else:
        if sh.cell_value(i, 67) < 2013:
            endPosition.text = str(sh.cell_value(i, 67)).strip('.0')
        else:
            endPositionTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 67), wb.datemode))
            endPositionList = list(map(int, endPositionTuple))
            endPositionYear = str(endPositionList[0])
            if endPositionList[1] < 10:
                endPositionMonth = '0' + str(endPositionList[1])
            else:
                endPositionMonth = str(endPositionList[1])
            if endPositionList[2] < 10:
                endPositionDay = '0' + str(endPositionList[2])
            else:
                endPositionDay = str(endPositionList[2])
            endPosition.text = endPositionYear + '-' + endPositionMonth + '-' + endPositionDay
    if endPosition.text == "" and beginPosition.text == "":
        EX_Extent.remove(temporalElement)
    distributionInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}distributionInfo')
    MD_Distribution = SubElement(distributionInfo, '{http://www.isotc211.org/2005/gmd}MD_Distribution')
    distributionFormat = SubElement(MD_Distribution, '{http://www.isotc211.org/2005/gmd}distributionFormat')
    MD_Format = SubElement(distributionFormat, '{http://www.isotc211.org/2005/gmd}MD_Format')
    name = SubElement(MD_Format, '{http://www.isotc211.org/2005/gmd}name')
    formatName = SubElement(name, '{http://www.isotc211.org/2005/gco}CharacterString')
    formatName.text = sh.cell_value(i, 68)
    version = SubElement(MD_Format, '{http://www.isotc211.org/2005/gmd}version')
    formatVersion = SubElement(version, '{http://www.isotc211.org/2005/gco}CharacterString')
    formatVersion.text = sh.cell_value(i, 69)
    transferOptions = SubElement(MD_Distribution, '{http://www.isotc211.org/2005/gmd}transferOptions')
    MD_DigitalTransferOptions = SubElement(transferOptions,
                                           '{http://www.isotc211.org/2005/gmd}MD_DigitalTransferOptions')
    onLine = SubElement(MD_DigitalTransferOptions, '{http://www.isotc211.org/2005/gmd}onLine')
    CI_OnlineResource = SubElement(onLine, '{http://www.isotc211.org/2005/gmd}CI_OnlineResource')
    linkage = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}linkage')
    URL = SubElement(linkage, '{http://www.isotc211.org/2005/gmd}URL')
    URL.text = sh.cell_value(i, 70)
    function = SubElement(CI_OnlineResource, '{http://www.isotc211.org/2005/gmd}function')
    CI_OnlineFunctionCode = SubElement(function, '{http://www.isotc211.org/2005/gmd}CI_OnLineFunctionCode', {
        'codeList': "http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_OnLineFunctionCode",
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 71)})
    CI_OnlineFunctionCode.text = CI_OnlineFunctionCode.get('codeListValue')
    dataQualityInfo = SubElement(root, '{http://www.isotc211.org/2005/gmd}dataQualityInfo')
    DQ_DataQuality = SubElement(dataQualityInfo, '{http://www.isotc211.org/2005/gmd}DQ_DataQuality')
    scope = SubElement(DQ_DataQuality, '{http://www.isotc211.org/2005/gmd}scope')
    DQ_Scope = SubElement(scope, '{http://www.isotc211.org/2005/gmd}DQ_Scope')
    level = SubElement(DQ_Scope, '{http://www.isotc211.org/2005/gmd}level')
    MD_ScopeCode = SubElement(level, '{http://www.isotc211.org/2005/gmd}MD_ScopeCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#MD_ScopeCode',
        'codeSpace': "ISOTC211/19115", 'codeListValue': sh.cell_value(i, 72)})
    MD_ScopeCode.text = MD_ScopeCode.get('codeListValue')
    report = SubElement(DQ_DataQuality, '{http://www.isotc211.org/2005/gmd}report')
    DQ_DomainConsistency = SubElement(report, '{http://www.isotc211.org/2005/gmd}DQ_DomainConsistency')
    result = SubElement(DQ_DomainConsistency, '{http://www.isotc211.org/2005/gmd}result')
    DQ_ConformanceResult = SubElement(result, '{http://www.isotc211.org/2005/gmd}DQ_ConformanceResult')
    specification = SubElement(DQ_ConformanceResult, '{http://www.isotc211.org/2005/gmd}specification')
    CI_Citation = SubElement(specification, '{http://www.isotc211.org/2005/gmd}CI_Citation')
    title = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}title')
    titleString = SubElement(title, '{http://www.isotc211.org/2005/gco}CharacterString')
    titleString.text = sh.cell_value(i, 73)
    date1 = SubElement(CI_Citation, '{http://www.isotc211.org/2005/gmd}date')
    CI_Date = SubElement(date1, '{http://www.isotc211.org/2005/gmd}CI_Date')
    date2 = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}date')
    reportDate = SubElement(date2, '{http://www.isotc211.org/2005/gco}Date')
    if sh.cell_value(i, 74) == "":
        reportDate.text = ""
    else:
        if sh.cell_value(i, 74) < 2013:
            reportDate.text = str(sh.cell_value(i, 74)).strip('.0')
        else:
            reportDateTuple = (xlrd.xldate_as_tuple(sh.cell_value(i, 74), wb.datemode))
            reportDateList = list(map(int, reportDateTuple))
            reportYear = str(reportDateList[0])
            if reportDateList[1] < 10:
                reportMonth = '0' + str(reportDateList[1])
            else:
                reportMonth = str(reportDateList[1])
            if reportDateList[2] < 10:
                reportDay = '0' + str(reportDateList[2])
            else:
                reportDay = str(reportDateList[2])
            reportDate.text = reportYear + '-' + reportMonth + '-' + reportDay
    dateType = SubElement(CI_Date, '{http://www.isotc211.org/2005/gmd}dateType')
    CI_DateTypeCode = SubElement(dateType, '{http://www.isotc211.org/2005/gmd}CI_DateTypeCode', {
        'codeList': 'http://www.isotc211.org/2005/resources/Codelist/gmxCodelists.xml#CI_DateTypeCode',
        'codeSpace': 'ISOTC211/19115', 'codeListValue': sh.cell_value(i, 75)})
    CI_DateTypeCode.text = CI_DateTypeCode.get('codeListValue')
    explanation = SubElement(DQ_ConformanceResult, '{http://www.isotc211.org/2005/gmd}explanation')
    explanationString = SubElement(explanation, '{http://www.isotc211.org/2005/gco}CharacterString')
    explanationString.text = sh.cell_value(i, 76)
    validationPerformed = SubElement(DQ_ConformanceResult, '{http://www.isotc211.org/2005/gmd}pass')
    validationString = SubElement(validationPerformed, '{http://www.isotc211.org/2005/gco}Boolean')
    validationString.text = str(sh.cell_value(i, 77))
    lineage = SubElement(DQ_DataQuality, '{http://www.isotc211.org/2005/gmd}lineage')
    LI_Lineage = SubElement(lineage, '{http://www.isotc211.org/2005/gmd}LI_Lineage')
    statement = SubElement(LI_Lineage, '{http://www.isotc211.org/2005/gmd}statement')
    statementString = SubElement(statement, '{http://www.isotc211.org/2005/gco}CharacterString')
    statementString.text = sh.cell_value(i, 78)
    # After all elements have been filled out, remove empty and non-required elements
    remove_graphic(MD_dataID, graphicOverview)
    remove_constraints(MD_dataID, resourceConstraints)
    remove_thesaurus(MD_dataID, descriptiveKeywords)
    remove_quality(root)
    tree = etree.ElementTree(root)
    #f = open("resource_%i.xml" % i, 'w')
    tree.write("resource_%i.xml" % i)  #, encoding="ISO-8859-1")
    print("Tree created for " + resourceTitle.text + "\n")