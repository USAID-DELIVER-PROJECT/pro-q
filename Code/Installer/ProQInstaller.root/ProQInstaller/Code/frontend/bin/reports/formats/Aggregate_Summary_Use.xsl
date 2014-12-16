<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="raw-xml.xsl"?>

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


	<TABLE border="1" cellspacing="0" width="80%">
	<TR>
	<TD width="40%" align = "center" bgcolor="#C0C0C0"><b>Required Funds</b></TD>
        <xsl:apply-templates select="Report/Other/Required_Funds" />		    
	</TR>
	<TR>
	<TD width="40%" align = "center" bgcolor="#C0C0C0"><b>Available Funds</b></TD>
        <xsl:apply-templates select="Report/Other/Available_Funds" />		    
	</TR>
	<TR>
	<TD width="40%" align = "center" bgcolor="#C0C0C0"><b>Ratio</b></TD>
        <xsl:apply-templates select="Report/Other/Ratio" />		    
	</TR>
	</TABLE>

	<p></p>	

	<TABLE border="1" cellspacing="0" width="100%">

	<TR>
	<TD bgcolor="#C0C0C0"> <b><u>Use</u></b> </TD>
	<TD bgcolor="#C0C0C0"> <b><u>Funding Source</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Amount</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Percent</u></b> </TD>	
	</TR>

	<xsl:for-each select="Report/Financials/SelectedBrand">
	  <TR>
          <xsl:apply-templates select="strUse" />
          <xsl:apply-templates select="strFundingSource" />      
          <xsl:apply-templates select="curAmount" />      	  
          <xsl:apply-templates select="sngPercent" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="strUse">
  <TD width="30%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="strFundingSource">
	<TD width="30%" > <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="curAmount">
	<TD width="20%" align="right"> <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="sngPercent">
	<TD width="20%" align="right"> <xsl:value-of /> </TD>
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

<xsl:template match="Report/Other/Required_Funds">
	<TD align = "right" width="40%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Available_Funds">
	<TD align = "right" width="40%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Ratio">
	<TD align = "right" width="40%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Storage_Budget">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Customs_Budget">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>


</xsl:stylesheet>