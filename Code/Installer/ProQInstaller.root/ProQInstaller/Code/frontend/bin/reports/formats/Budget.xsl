<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="Budget.xsl"?>

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
	<TD bgcolor="#C0C0C0"> <b><u>Test Kit</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Kits Required</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Cost of Required</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Qty to Procure</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Cost of Procured</u></b> </TD>
	</TR>

	<xsl:for-each select="Report/Financials/SelectedBrand">
	  <TR>
          <xsl:apply-templates select="Brand" />
          <xsl:apply-templates select="Kits_Req" />      
          <xsl:apply-templates select="Costs_Req" />      	  
          <xsl:apply-templates select="Qty_Proq" />      
          <xsl:apply-templates select="Cost_Proq" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>

	<TABLE border="1" cellspacing="0" width="100%">
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Storage/Distribution Budget</b></TD>
        <xsl:apply-templates select="Report/Other/Storage_Budget" />		    
	</TR>
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Customs Budget</b></TD>
        <xsl:apply-templates select="Report/Other/Customs_Budget" />		    
	</TR>
	</TABLE>

	<p></p>

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="Brand">
  <TD width="25%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="Kits_Req">
	<TD width="18%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Costs_Req">
	<TD width="18%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Qty_Proq">
	<TD width="18%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Cost_Proq">
	<TD width="21%" align="right"> &#160; <xsl:value-of /> </TD>
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
	<TD align = "right" width="40%"> &#160; <xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Available_Funds">
	<TD align = "right" width="40%"> &#160; <xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Ratio">
	<TD align = "right" width="40%"> &#160; <xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Storage_Budget">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Customs_Budget">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>


</xsl:stylesheet>