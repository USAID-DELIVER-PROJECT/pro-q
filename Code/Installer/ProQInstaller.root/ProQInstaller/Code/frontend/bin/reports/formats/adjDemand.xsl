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
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Demand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Service Capacity</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Demand Adjusted for Service Capacity</u></b> </TD>	
	</TR>

	<xsl:for-each select="Report/Demands/Brand">
	  <TR>
          <xsl:apply-templates select="Name" />
          <xsl:apply-templates select="Demand" />      
          <xsl:apply-templates select="Service_Capacity" />      
          <xsl:apply-templates select="SC_Adjusted_Demand" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>
       
	<TABLE border="1" cellspacing="0" width="100%">

	<TR>
	<TD bgcolor="#C0C0C0"> <b><u>Brand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>SC Adj. Demand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Quality Control</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Wastage</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Adj. Demand</u></b> </TD>
	</TR>

	<xsl:for-each select="Report/Demands/Brand">
	  <TR>
          <xsl:apply-templates select="Name" />
          <xsl:apply-templates select="SC_Adjusted_Demand" />      
          <xsl:apply-templates select="Quality_Control" />      
          <xsl:apply-templates select="Wastage" />      
          <xsl:apply-templates select="Adjusted_Demand" />        
	  </TR>
	</xsl:for-each> 
	</TABLE>


    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="Name">
  <TD width="25%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="Demand">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Service_Capacity">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="SC_Adjusted_Demand">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Quality_Control">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Wastage">
	<TD width="15%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Adjusted_Demand">
	<TD width="20%" align="right"> &#160; <xsl:value-of /> </TD>
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



</xsl:stylesheet>