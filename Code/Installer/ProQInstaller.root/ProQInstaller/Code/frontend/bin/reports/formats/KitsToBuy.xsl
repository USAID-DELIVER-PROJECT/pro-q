<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="KitstoBuy.xsl"?>

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
	

	<H3>Adjusted Demand</H3>	
	<TABLE border="1" cellspacing="0" width="100%">

	<xsl:for-each select="Report/Adjusted_Demand/AD">
	  <TR>
          <xsl:apply-templates select="Name" />
          <xsl:apply-templates select="Value" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>


	<H3>Quantity Required</H3>	       
	<TABLE border="1" cellspacing="0" width="100%">
	<xsl:for-each select="Report/Qty_Required/AD">
	  <TR>
          <xsl:apply-templates select="Name" />
          <xsl:apply-templates select="Value" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>

	<H3>Financial Requirements</H3>	       
	<TABLE border="1" cellspacing="0" width="100%">
	<xsl:for-each select="Report/Financial_Requirements/AD">
	  <TR>
          <xsl:apply-templates select="Name" />
          <xsl:apply-templates select="Value" />      
	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>


	<H3>Budget Reconciliation</H3>	       
	<TABLE border="1" cellspacing="0" width="100%">
	  <TR>
            <TD bgcolor="#C0C0C0" > &#160;  </TD>
	    <TD align = "center" bgcolor="#C0C0C0" >Funds</TD>
	    <TD align = "center" bgcolor="#C0C0C0" >Quantity</TD>	
	    <TD align = "center" bgcolor="#C0C0C0" >Cost</TD>	
	  </TR>

	<xsl:for-each select="Report/Budget/AD">
	  <TR>
          <xsl:apply-templates select="Type" />
          <xsl:apply-templates select="Funds" />      
          <xsl:apply-templates select="Qty" />
          <xsl:apply-templates select="Cost" />      

	  </TR>
	</xsl:for-each> 
	</TABLE>

	<p></p>


        <H3>Quantity to Procure</H3>	       
	<TABLE border="1" cellspacing="0" width="100%">
	  <TR>
            <TD bgcolor="#C0C0C0" > &#160;  </TD>
            <TD bgcolor="#C0C0C0" > &#160;  </TD>
	    <TD align = "center" bgcolor="#C0C0C0" >Cost</TD>
	    <TD align = "center" bgcolor="#C0C0C0" >Quantity</TD>	
	  </TR>

	<xsl:for-each select="Report/Procure/AD">
	  <TR>	 
          <xsl:apply-templates select="Selected" />
          <xsl:apply-templates select="ProcType" />      
          <xsl:apply-templates select="ProCost" />
          <xsl:apply-templates select="ProQty" />      

	  </TR>
	</xsl:for-each> 
	</TABLE>



    </BODY>
  </HTML>
</xsl:template>

<xsl:template match="Name">
  <TD width="50%" >
	<b> &#160; <xsl:value-of /> </b>
  </TD>
</xsl:template>

<xsl:template match="Value">
  <TD width="50%" align = "right">
	&#160; <xsl:value-of />
  </TD>
</xsl:template>

<xsl:template match="Type">
	<TD width="25%" align="Left"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Funds">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Qty">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Cost">
	<TD width="25%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="Selected">
	<TD width="10%" align="Center"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="ProcType">
	<TD width="30%" align="Left"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="ProQty">
	<TD width="30%" align="right"> &#160; <xsl:value-of /> </TD>
</xsl:template>

<xsl:template match="ProCost">
	<TD width="30%" align="right"> &#160; <xsl:value-of /> </TD>
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