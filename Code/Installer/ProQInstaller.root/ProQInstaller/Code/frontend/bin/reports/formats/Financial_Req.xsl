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
	<TD bgcolor="#C0C0C0"> <b><u>Brand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Tests Required</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Tests per Kit</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Kits Required</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Cost per Kit</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Line Total</u></b> </TD>		
	</TR>

	<xsl:for-each select="Report/Financials/SelectedBrand">
	  <TR>
          <xsl:apply-templates select="Brand" />
          <xsl:apply-templates select="Tests_req" />
          <xsl:apply-templates select="Tests_Kit" />      
          <xsl:apply-templates select="Kits_req" />      
          <xsl:apply-templates select="Costs_kit" />      	  
          <xsl:apply-templates select="Line_Total" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>

	<TABLE border="1" cellspacing="0" width="100%">
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Total Kit Costs</b></TD>
        <xsl:apply-templates select="Report/Other/Kit_Cost" />		    
	</TR>
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Storage/Distribution Costs</b></TD>
        <xsl:apply-templates select="Report/Other/Storage_Cost" />		    
	</TR>
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Customs Costs</b></TD>
        <xsl:apply-templates select="Report/Other/Customs_Cost" />		    
	</TR>
	<TR>
	<TD width="85%" align = "center" bgcolor="#C0C0C0"><b>Grand Total</b></TD>
        <xsl:apply-templates select="Report/Other/Total_Cost" />		    
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

<xsl:template match="Tests_req">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Tests_Kit">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Kits_req">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Costs_kit">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Line_Total">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
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

<xsl:template match="Report/Other/Kit_Cost">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Customs_Cost">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Storage_Cost">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>

<xsl:template match="Report/Other/Total_Cost">
	<TD align = "right" width="15%"><xsl:value-of /></TD>	
</xsl:template>


</xsl:stylesheet>