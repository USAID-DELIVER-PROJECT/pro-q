<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="Required.xsl"?>

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
	

        <!-- Column Headers -->
	<TABLE border="1" cellspacing="0" width="100%">
        
	<TR>
	<TD bgcolor="#C0C0C0"> <b><u>Brand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Adjusted Demand</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Lead Time Stock</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Buffer Stock</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Qty On Hand</u></b> </TD>		
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Qty On Order</u></b> </TD>		
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Qty Required</u></b> </TD>
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Shipment Frequency</u></b> </TD>	
	<TD align = "center" bgcolor="#C0C0C0"> <b><u>Volume / Shipment (cu. m)</u></b> </TD>			
	</TR>

	<xsl:for-each select="Report/Financials/SelectedBrand">
	  <TR>
          <xsl:apply-templates select="Brand" />
          <xsl:apply-templates select="Adj_Demand" />
          <xsl:apply-templates select="Lead_Time_Stock" />      
          <xsl:apply-templates select="Buffer_Stock" />      
          <xsl:apply-templates select="Qty_Hand" />      
          <xsl:apply-templates select="Qty_Order" />      
          <xsl:apply-templates select="Qty_Required" />  
          <xsl:apply-templates select="Ship_Frequency" />      
          <xsl:apply-templates select="Ship_Volume" />          
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>

    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="Brand">
  <TD width="21%" >
	<b> <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="Adj_Demand">
	<TD width="10%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Lead_Time_Stock">
	<TD width="9%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Buffer_Stock">
	<TD width="9%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Ship_Frequency">
	<TD width="10%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Storage_Capacity">
	<TD width="10%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Qty_Hand">
	<TD width="10%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Qty_Order">
	<TD width="10%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Qty_Required">
	<TD width="11%" align="right"> &#160; <b> <xsl:value-of />  </b> </TD>
</xsl:template>

<xsl:template match="Ship_Volume">
	<TD width="10%" align="right"> &#160; <b> <xsl:value-of /> </b> </TD>
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