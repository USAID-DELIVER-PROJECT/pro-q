<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="Demand.xsl"?>

<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">

<xsl:template match="/">
  <HTML>	
    <BODY>

	<table border="0" width="100%" cellspacing="0">
	  <tr>
    	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Aggregation" />		    
	  </tr>
	  <tr>
	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Program" />		    
	  </tr>
	  <tr>
	    <td width="70%"></td>
	    <xsl:apply-templates select="Report/Header/Use" />		    
	  </tr>
	</table>

        <xsl:apply-templates select="Report/Header/Report_Title" />		    
	

	<TABLE border="1" cellspacing="0" width="100%">

	  <TR>
	   <TD bgcolor="#C0C0C0"> <b><u>Methodology</u></b> </TD>
	   <TD align = "center" bgcolor="#C0C0C0"> <b><u>Brand</u></b> </TD>
	   <TD align = "center" bgcolor="#C0C0C0"> <b><u>Demand</u></b> </TD>	
	  </TR>


	  <!-- IF any Logistics Methodology then Display row. -->
	  <xsl:if test="Report/Methodologies/MethodBrand[$any$ strMethod='Logistics']">	
	  <TR>	
	    <TD>Logistics</TD>	
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Logistics']">
                <TR>
                <xsl:apply-templates select="strBrand" />                      
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>    
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Logistics']">
                <TR>
                 <xsl:apply-templates select="varDemand" />
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>
	  </TR>
          </xsl:if>	    

          <xsl:if test="Report/Methodologies/MethodBrand[$any$ strMethod='Demographic/Morbidity']">
	  <TR>
	    <TD>Demographic/Morbidity</TD>	
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Demographic/Morbidity']">
                <TR>
                <xsl:apply-templates select="strBrand" />                      
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>    
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Demographic/Morbidity']">
                <TR>
                 <xsl:apply-templates select="varDemand" />
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>
	  </TR>	    
          </xsl:if>

          <xsl:if test="Report/Methodologies/MethodBrand[$any$ strMethod='Service Statistics']">
	  <TR>
	    <TD>Service Statistics</TD>	
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Service Statistics']">
                <TR>
                <xsl:apply-templates select="strBrand" />                      
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>    
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Service Statistics']">
                <TR>
                 <xsl:apply-templates select="varDemand" />
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>
	  </TR>	    
          </xsl:if>
	
          <xsl:if test="Report/Methodologies/MethodBrand[$any$ strMethod='Target']">
	  <TR>
	    <TD>Target</TD>	
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Target']">
                <TR>
                <xsl:apply-templates select="strBrand" />                      
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>    
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Target']">
                <TR>
                 <xsl:apply-templates select="varDemand" />
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>
	  </TR>	    
          </xsl:if> 

          <xsl:if test="Report/Methodologies/MethodBrand[$any$ strMethod='Average']">
	  <TR>
	    <TD>Average</TD>	
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Average']">
                <TR>
                <xsl:apply-templates select="strBrand" />                      
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>    
	    <TD>
            <TABLE border="1" cellspacing="0" width="100%">
	      <xsl:for-each select="Report/Methodologies/MethodBrand[strMethod = 'Average']">
                <TR>
                 <xsl:apply-templates select="rs:forcenull" /> 
                 <xsl:apply-templates select="varDemand" />
                </TR>
              </xsl:for-each>
            </TABLE>
            </TD>
	  </TR>	    
          </xsl:if> 

	</TABLE>

	<p></p>

	<TABLE border="1" cellspacing="0" width="80%">
	<TR>
	<TD width="40%" align = "center" bgcolor="#C0C0C0"><b>Selected Methodology</b></TD>
        <xsl:apply-templates select="Report/Other/Selected_Method" />		    
	</TR>
	</TABLE>

	<p></p>

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="strMethod">
  <TD width="40%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="strBrand">
	<TD width="40%"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="varDemand">
	<TD width="20%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="rs:forcenull">
	<TD width="20%" align="right"> &#160;</TD>
</xsl:template>

<xsl:template match="Report/Header/Aggregation">
	<td width="30%"><b><xsl:value-of /></b></td>	
</xsl:template>

<xsl:template match="Report/Header/Program">
	<td width="30%"><i><xsl:value-of /></i></td>	
</xsl:template>

<xsl:template match="Report/Header/Use">
	<td width="30%"><i><xsl:value-of /></i></td>	
</xsl:template>

<xsl:template match="Report/Header/Report_Title">
	<H2><xsl:value-of /></H2>	
</xsl:template>

<xsl:template match="Report/Other/Selected_Method">
	<TD align = "center" width="40%"><xsl:value-of /></TD>	
</xsl:template>


</xsl:stylesheet>