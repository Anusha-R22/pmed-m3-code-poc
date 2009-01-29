<?xml version="1.0" encoding="ISO-8859-1" ?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'
xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'
xmlns:rs='urn:schemas-microsoft-com:rowset'
xmlns:z='#RowsetSchema'>
  <xsl:output method="xml" version="1.0"  encoding="ISO-8859-1"/>
  <xsl:key name="visit-search"
  match="//root/StudyVisit/xml/rs:data/z:row" use="@VISITID" />

  <xsl:key name="studyvisitcrfpage-search"
  match="//root/StudyVisitCRFPage/xml/rs:data/z:row"
  use="@VISITID" />

  <xsl:key name="crfpage-search"
  match="//root/CRFPage/xml/rs:data/z:row" use="@CRFPAGEID" />

  <xsl:key name="crfelement-search"
  match="//root/CRFElement/xml/rs:data/z:row" use="@CRFPAGEID" />

  <xsl:key name="dataitem-search"
  match="//root/DataItem/xml/rs:data/z:row" use="@DATAITEMID" />

  <xsl:key name="valuedata-search"
  match="//root/ValueData/xml/rs:data/z:row" use="@DATAITEMID" />

<xsl:key name="eformgroup-search"
  match="//root/PageGroups/xml/rs:data/z:row" use="@CRFPAGEID" />
  
  <xsl:template match="/">
    <xsl:element name="ODM">
      <xsl:attribute name="FileType">Snapshot</xsl:attribute>

      <xsl:attribute name="FileOID">Test</xsl:attribute>

      <xsl:attribute name="CreationDateTime">
        <xsl:value-of select="//root/@created" />
      </xsl:attribute>

      <xsl:apply-templates />
    </xsl:element>
  </xsl:template>

  <xsl:template match="ClinicalTrial/xml/rs:data/z:row">
    <xsl:element name="Study">
      <xsl:attribute name="OID">
        <xsl:value-of select="@CLINICALTRIALNAME" />
      </xsl:attribute>

      <xsl:element name="GlobalVariables">
        <xsl:element name="StudyName">
          <xsl:value-of select="@CLINICALTRIALNAME" />
        </xsl:element>

        <xsl:element name="StudyDescription">
          <xsl:value-of select="@CLINICALTRIALDESCRIPTION" />
        </xsl:element>

        <xsl:element name="ProtocolName">
          <xsl:value-of select="@CLINICALTRIALNAME" />
        </xsl:element>				
				

      </xsl:element>

      <xsl:element name="BasicDefinitions" />

      <xsl:element name="MetaDataVersion">
        <xsl:attribute name="OID">1</xsl:attribute>

        <xsl:attribute name="Name">Version 1</xsl:attribute>

        <xsl:element name="Protocol">
          <xsl:for-each
          select="//root/StudyVisit/xml/rs:data/z:row">
            <xsl:element name="StudyEventRef">
              <xsl:attribute name="StudyEventOID">
                <xsl:value-of select="@VISITCODE" />
              </xsl:attribute>

              <xsl:attribute name="OrderNumber">
                <xsl:value-of select="@VISITORDER" />
              </xsl:attribute>

              <xsl:attribute name="Mandatory">Yes</xsl:attribute>
            </xsl:element>
          </xsl:for-each>
        </xsl:element>

        <xsl:for-each
        select="//root/StudyVisit/xml/rs:data/z:row">
          <xsl:element name="StudyEventDef">
            <xsl:attribute name="OID">
              <xsl:value-of select="@VISITCODE" />
            </xsl:attribute>

            <xsl:attribute name="Name">
              <xsl:value-of select="@VISITNAME" />
            </xsl:attribute>

            <xsl:attribute name="Repeating">
				      	<xsl:choose>
							<xsl:when test="@REPEATING">Yes</xsl:when>
							<xsl:otherwise>No</xsl:otherwise>
						</xsl:choose>
            </xsl:attribute>

			<xsl:attribute name="Type">Scheduled</xsl:attribute>
			
            <xsl:for-each
            select="key('studyvisitcrfpage-search', @VISITID)">
              <xsl:element name="FormRef">
                <xsl:attribute name="FormOID">
                  <xsl:value-of
                  select="key('crfpage-search', @CRFPAGEID)/attribute::CRFPAGECODE" />
                </xsl:attribute>

                <xsl:attribute name="OrderNumber">
                  <xsl:value-of
                  select="key('crfpage-search', @CRFPAGEID)/attribute::CRFPAGEORDER" />
                </xsl:attribute>

                <xsl:attribute name="Mandatory">Yes</xsl:attribute>
              </xsl:element>
            </xsl:for-each>
          </xsl:element>
        </xsl:for-each>

        <xsl:for-each select="//root/CRFPage/xml/rs:data/z:row">
          <xsl:element name="FormDef">
            <xsl:attribute name="OID">
              <xsl:value-of select="@CRFPAGECODE" />
            </xsl:attribute>

            <xsl:attribute name="Name">
              <xsl:value-of select="@CRFTITLE" />
            </xsl:attribute>

            <xsl:attribute name="Repeating">No</xsl:attribute>

            <xsl:for-each
            select="key('eformgroup-search', @CRFPAGEID)">

				<xsl:element name="ItemGroupRef">
				  <xsl:attribute name="ItemGroupOID">
					<xsl:value-of select="@QGROUPCODE" />
				  </xsl:attribute>
	
				  <xsl:attribute name="OrderNumber">1</xsl:attribute>
	
				  <xsl:attribute name="Mandatory">Yes</xsl:attribute>
				</xsl:element>
			</xsl:for-each>
          </xsl:element>
        </xsl:for-each>
	
        <xsl:for-each select="//root/PageGroups/xml/rs:data/z:row">
          <xsl:element name="ItemGroupDef">
			<xsl:attribute name="OID">
              <xsl:value-of select="@QGROUPCODE" />
            </xsl:attribute>

            <xsl:attribute name="Name">
              <xsl:value-of select="@QGROUPNAME" />
            </xsl:attribute>

            <xsl:attribute name="Repeating">
			  <xsl:choose>
				<xsl:when test="@OWNERQGROUPID = 0">No</xsl:when>
				<xsl:otherwise>Yes</xsl:otherwise>
			  </xsl:choose>
            </xsl:attribute>

			<xsl:variable name="var_OwnerQGroupId" select="@OWNERQGROUPID" />
			<xsl:variable name="var_CRFPageId" select="@CRFPAGEID" />

			<xsl:for-each select="//root/CRFElement/xml/rs:data/z:row[attribute::CRFPAGEID = $var_CRFPageId and attribute::OWNERQGROUPID = $var_OwnerQGroupId]">
	          <xsl:element name="ItemRef">
				<xsl:attribute name="ItemOID">
				  <xsl:value-of select="@DATAITEMCODE" />
				</xsl:attribute>
				
				<xsl:variable name="var_DataItemId" select="@DATAITEMID" />

				<xsl:attribute name="OrderNumber">
				  <xsl:choose>
					<xsl:when test="@OWNERQGROUPID = 0">
						<xsl:value-of select="@FIELDORDER" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="//root/RQGDetail/xml/rs:data/z:row[attribute::CRFPAGEID = $var_CRFPageId and attribute::DATAITEMID = $var_DataItemId and attribute::QGROUPID = $var_OwnerQGroupId]">
							<xsl:value-of select="@QORDER" />
						</xsl:for-each>
					</xsl:otherwise>
				  </xsl:choose>
				</xsl:attribute>

				<xsl:attribute name="Mandatory">
				  <xsl:choose>
					<xsl:when test="@OWNERQGROUPID = 0">
					  <xsl:choose>
						<xsl:when test="@MANDATORY = 0">No</xsl:when>
						<xsl:otherwise>Yes</xsl:otherwise>
					  </xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="//root/RQGDetail/xml/rs:data/z:row[attribute::CRFPAGEID = $var_CRFPageId and attribute::DATAITEMID = $var_DataItemId and attribute::QGROUPID = $var_OwnerQGroupId]">
						  <xsl:choose>
							<xsl:when test="@MANDATORY = 0">No</xsl:when>
							<xsl:when test="@MANDATORY = 1">Yes</xsl:when>
						  </xsl:choose>
						</xsl:for-each>
					</xsl:otherwise>
				  </xsl:choose>
				</xsl:attribute>
			  </xsl:element>
			  
			</xsl:for-each>

		  </xsl:element>
		  </xsl:for-each>
	
        <xsl:for-each select="//root/DataItem/xml/rs:data/z:row">
          <xsl:element name="ItemDef">
            <xsl:attribute name="OID">
              <xsl:value-of select="@DATAITEMCODE" />
            </xsl:attribute>

            <xsl:attribute name="Name">
              <xsl:value-of select="@DATAITEMNAME" />
            </xsl:attribute>

            <xsl:attribute name="DataType">
              <xsl:choose>
                <xsl:when test="@DATATYPE = 0">text</xsl:when>

                <xsl:when test="@DATATYPE = 1">text</xsl:when>

                <xsl:when test="@DATATYPE = 2">integer</xsl:when>

                <xsl:when test="@DATATYPE = 3">float</xsl:when>

                <xsl:when test="@DATATYPE = 4">
                datetime</xsl:when>

                <xsl:when test="@DATATYPE = 5">
                multimedia</xsl:when>

                <xsl:when test="@DATATYPE = 6">float</xsl:when>
              </xsl:choose>
            </xsl:attribute>

            <xsl:attribute name="Length">
              <xsl:value-of select="@DATAITEMLENGTH" />
            </xsl:attribute>

            <xsl:attribute name="SASFieldName">
              <xsl:value-of select="@EXPORTNAME" />
            </xsl:attribute>
          </xsl:element>
        </xsl:for-each>

        <xsl:for-each
        select="//root/CodeLists/xml/rs:data/z:row">
          <xsl:element name="CodeList">
            <xsl:attribute name="OID">
              <xsl:value-of select="@DATAITEMID" />
            </xsl:attribute>

            <xsl:attribute name="Name">
              <xsl:value-of select="@DATAITEMCODE" />
            </xsl:attribute>

            <xsl:attribute name="DataType">text</xsl:attribute>

            <xsl:for-each
            select="key('valuedata-search', @DATAITEMID)">
              <xsl:element name="CodeListItem">
                <xsl:attribute name="CodedValue">
                  <xsl:value-of select="@VALUECODE" />
                </xsl:attribute>

                <xsl:element name="Decode">
                  <xsl:element name="TranslatedText">
                    <xsl:attribute name="xml:lang">
                    en</xsl:attribute>

                    <xsl:value-of select="@ITEMVALUE" />
                  </xsl:element>
                </xsl:element>
              </xsl:element>
            </xsl:for-each>
          </xsl:element>
        </xsl:for-each>
      </xsl:element>
    </xsl:element>

    <xsl:element name="ClinicalData">
      <xsl:attribute name="StudyOID">
        <xsl:value-of select="@CLINICALTRIALNAME" />
      </xsl:attribute>

      <xsl:attribute name="MetaDataVersionOID">1</xsl:attribute>

      <xsl:for-each
      select="//root/TrialSubject/xml/rs:data/z:row">
        <xsl:element name="SubjectData">
          <xsl:attribute name="SubjectKey">
            <xsl:value-of select="@LOCALIDENTIFIER1" />
          </xsl:attribute>

          <xsl:variable name="var_PersonId" select="@PERSONID" />

          <xsl:for-each
          select="//root/VisitInstance/xml/rs:data/z:row[attribute::PERSONID = $var_PersonId]">

            <xsl:element name="StudyEventData">
              <xsl:attribute name="StudyEventOID">
                <xsl:value-of
                select="key('visit-search', @VISITID)/attribute::VISITCODE" />
              </xsl:attribute>
          <xsl:attribute name="StudyEventRepeatKey">
            <xsl:value-of select="@VISITCYCLENUMBER" />
          </xsl:attribute>
              <xsl:variable name="var_VisitId"
              select="@VISITID" />

              <xsl:variable name="var_VisitCycleNumber"
              select="@VISITCYCLENUMBER" />

              <xsl:for-each
              select="//root/CRFPageInstance/xml/rs:data/z:row[attribute::VISITID = $var_VisitId][attribute::VISITCYCLENUMBER = $var_VisitCycleNumber][attribute::PERSONID = $var_PersonId] ">

                <xsl:element name="FormData">
                  <xsl:attribute name="FormOID">
                    <xsl:value-of
                    select="key('crfpage-search', @CRFPAGEID)/attribute::CRFPAGECODE" />
                  </xsl:attribute>
					<xsl:attribute name="FormRepeatKey">
						<xsl:value-of select="@CRFPAGECYCLENUMBER" />
					</xsl:attribute>
					<xsl:variable name="var_CRFPageTaskId" select="@CRFPAGETASKID" />

				  <xsl:variable name="var_PageId" select="@CRFPAGEID" />

			      <xsl:for-each select="//root/PageGroups/xml/rs:data/z:row[attribute::CRFPAGEID = $var_PageId]">

					<xsl:variable name="var_OwnerQGroupId" select="@OWNERQGROUPID" />
					<xsl:variable name="var_QGroupCode" select="@QGROUPCODE" />

						<xsl:for-each
							select="//root/ResponseGroupInfo/xml/rs:data/z:row[attribute::PERSONID = $var_PersonId][attribute::CRFPAGETASKID = $var_CRFPageTaskId][attribute::QGROUPID = $var_OwnerQGroupId]">
							
							<xsl:variable name="var_QGroupRepeat" select="@REPEATNUMBER" />

							<xsl:element name="ItemGroupData">
								<xsl:attribute name="ItemGroupOID">
									<xsl:value-of select="$var_QGroupCode" />
								</xsl:attribute>
								<xsl:attribute name="ItemGroupRepeatKey">
									<xsl:value-of select="$var_QGroupRepeat" />
								</xsl:attribute>

								<xsl:for-each select="//root/CRFElement/xml/rs:data/z:row[attribute::CRFPAGEID = $var_PageId and attribute::OWNERQGROUPID = $var_OwnerQGroupId]">

									<xsl:variable name="var_DataItemId" select="@DATAITEMID" />

									<xsl:for-each
									select="//root/DataItemResponse/xml/rs:data/z:row[attribute::CRFPAGETASKID = $var_CRFPageTaskId and attribute::PERSONID = $var_PersonId and attribute::DATAITEMID = $var_DataItemId and attribute::REPEATNUMBER = $var_QGroupRepeat]">

									<xsl:element name="ItemData">
										<xsl:attribute name="ItemOID">
										<xsl:value-of
										select="key('dataitem-search', @DATAITEMID)/attribute::DATAITEMCODE" />
										</xsl:attribute>

										<xsl:attribute name="Value">
										<xsl:value-of
										select="@RESPONSEVALUE" />
										</xsl:attribute>
									</xsl:element>
									</xsl:for-each>
									
								</xsl:for-each>
								
							</xsl:element>
							
						</xsl:for-each>
						
					</xsl:for-each>
					
                </xsl:element>
              </xsl:for-each>
            </xsl:element>
          </xsl:for-each>
        </xsl:element>
      </xsl:for-each>
    </xsl:element>
  </xsl:template>
</xsl:stylesheet>

